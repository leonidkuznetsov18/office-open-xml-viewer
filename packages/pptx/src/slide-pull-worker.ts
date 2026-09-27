import { OoxmlResourceLimitError, type OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  HARD_MAX_PPTX_SLIDE_JSON_BYTES,
  exactTransferableArrayBuffer,
  PULL_SESSION_PROTOCOL,
  PullSessionHost,
  PullSessionHostCoordinator,
  normalizePullSessionInsufficientCreditError,
  deserializeWorkerError,
  parseResourceLimitError,
  serializeWorkerError,
  type PullSessionCommand,
  type PullSessionIdentity,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import type { Slide } from './types.js';
import {
  acknowledgePptxSlide,
  readPptxSlideCursorUsage,
  type PptxSlideAcceptor,
  type PptxSlideArchiveExecutor,
  type PptxSlideCursorArchive,
} from './slide-cursor-operation.js';
type LifecycleState = 'ready' | 'resetting' | 'reset-failed';
const RESETTING_ERROR_CODE = 'ooxml-pull-resetting';
const RESET_FAILED_ERROR_CODE = 'ooxml-pull-reset-failed';

/**
 * Format-owned adapter between PPTX's indivisible Rust slide unit and the
 * common pull-session ownership protocol. The presentation/viewer integration
 * intentionally lives elsewhere; this class owns only package serialization,
 * transfer, acknowledgement, and rollback.
 */
export class SlidePullWorker {
  private coordinatorGeneration = new PullSessionHostCoordinator();
  private readonly sessions = new Map<number, {
    host: PullSessionHost<ArrayBuffer, number>;
    identity: PullSessionIdentity<number>;
  }>();
  private readonly pendingOpens = new Map<number, {
    identity: PullSessionIdentity<number>;
    canceled: boolean;
  }>();
  private operationTail: Promise<void> = Promise.resolve();
  private resourceFailure: OoxmlResourceLimitError | undefined;
  private lifecycleState: LifecycleState = 'ready';
  private resetBarrier: Promise<void> | undefined;
  private resetIdentities = new Map<number, PullSessionIdentity<number>>();

  constructor(
    private readonly archive: () => PptxSlideCursorArchive | null | undefined,
    private readonly acceptSlide?: PptxSlideAcceptor,
    private readonly executeArchive: PptxSlideArchiveExecutor =
      (operation) => operation(this.requireArchive()),
  ) {}

  get coordinator(): PullSessionHostCoordinator {
    return this.coordinatorGeneration;
  }

  /** Register synchronously before a worker message handler's first await. */
  reserveOpen(identity: PullSessionIdentity<number>): void {
    this.assertReady();
    assertPositiveIdentity(identity);
    if (this.pendingOpens.has(identity.sessionId) || this.sessions.has(identity.sessionId)) {
      throw new Error('slide pull session id is already reserved');
    }
    this.pendingOpens.set(identity.sessionId, { identity, canceled: false });
  }

  abandonOpen(sessionId: number): void {
    this.pendingOpens.delete(sessionId);
  }

  get pendingOpenCount(): number {
    return this.pendingOpens.size;
  }

  async open(slideIndex: number, identity: PullSessionIdentity<number>): Promise<void> {
    this.assertReady();
    if (this.resourceFailure) throw this.resourceFailure;
    if (!Number.isSafeInteger(slideIndex) || slideIndex < 0) {
      throw new RangeError('slide index must be a non-negative safe integer');
    }
    const pending = this.pendingOpens.get(identity.sessionId);
    if (!pending || !sameIdentity(pending.identity, identity)) {
      throw new Error('slide pull session open reservation is stale or missing');
    }

    let completeOperation!: () => void;
    const completion = new Promise<void>((resolve) => { completeOperation = resolve; });
    const started = this.operationTail.then(() => this.coordinator.enqueue(async () => {
      if (pending.canceled) throw new Error('slide pull session open was canceled');
      let preparedSlide: Slide | undefined;
      let terminalPending = false;
      const host = new PullSessionHost<ArrayBuffer, number>({
        ...identity,
        maxByteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
        coordinator: this.coordinator,
        driver: {
          pull: (byteCredit) => {
            let bytes: Uint8Array;
            try {
              bytes = this.executeArchive((archive) => archive.pull_slide(
                slideIndex,
                identity.operationId,
                identity.generation,
                byteCredit,
              ));
            } catch (error) {
              const insufficient = normalizePullSessionInsufficientCreditError(
                error,
                byteCredit,
                HARD_MAX_PPTX_SLIDE_JSON_BYTES,
              );
              if (insufficient) throw insufficient;
              this.latchResourceFailure(error);
              throw error;
            }
            const payload = exactTransferableArrayBuffer(bytes);
            if (this.acceptSlide) {
              preparedSlide = JSON.parse(new TextDecoder().decode(new Uint8Array(payload))) as Slide;
            }
            terminalPending = true;
            return {
              payload,
              byteLength: payload.byteLength,
              done: true,
              transfer: [payload],
            };
          },
          measureChunk: ({ payload }) => payload.byteLength,
          acknowledge: () => {
            if (!terminalPending) throw new Error('slide unit is not awaiting acknowledgement');
            try {
              acknowledgePptxSlide(
                this.executeArchive,
                identity,
                slideIndex,
                preparedSlide,
                this.acceptSlide,
              );
            } catch (error) {
              this.latchResourceFailure(error);
              throw error;
            }
            terminalPending = false;
            preparedSlide = undefined;
            this.sessions.delete(identity.sessionId);
            completeOperation();
          },
          cancel: async () => {
            try {
              if (this.archive()) {
                await this.executeArchive((archive) => archive.cancel_slide());
              }
            } finally {
              preparedSlide = undefined;
              terminalPending = false;
              this.sessions.delete(identity.sessionId);
              completeOperation();
            }
          },
          // A protocol close ends this slide operation. The presentation-level
          // cursor remains reusable; reset() owns close_presentation_session().
          close: async () => {
            try {
              if (this.archive()) {
                await this.executeArchive((archive) => archive.cancel_slide());
              }
            } finally {
              preparedSlide = undefined;
              terminalPending = false;
              this.sessions.delete(identity.sessionId);
              completeOperation();
            }
          },
          resourceUsage: () => {
            try {
              return this.readResourceUsage();
            } catch (error) {
              this.latchResourceFailure(error);
              throw error;
            }
          },
        },
      });
      this.sessions.set(identity.sessionId, { host, identity });
      this.pendingOpens.delete(identity.sessionId);
    }));
    this.operationTail = started.then(() => completion, () => undefined);
    try {
      await started;
    } catch (error) {
      this.pendingOpens.delete(identity.sessionId);
      completeOperation();
      throw error;
    }
  }

  async postOpenedSafely(
    identity: PullSessionIdentity<number>,
    postOpened: () => void,
    postError: (error: unknown) => void,
  ): Promise<void> {
    if (this.lifecycleState !== 'ready') {
      try {
        postError(this.lifecycleError());
      } catch {
        // Reset owns cleanup and the response channel is unavailable.
      }
      return;
    }
    try {
      postOpened();
    } catch (error) {
      await this.closeIdentity(identity);
      try {
        postError(error);
      } catch {
        // The response channel is gone; the package operation is already clean.
      }
    }
  }

  dispatch(
    command: PullSessionCommand<number>,
    post: (response: PullSessionResponse<ArrayBuffer, number>, transfer?: Transferable[]) => void,
  ): Promise<void> {
    if (this.lifecycleState !== 'ready') {
      post(this.responseDuringReset(command));
      return Promise.resolve();
    }
    const session = this.sessions.get(command.sessionId);
    if (session) return session.host.dispatch(command, post);
    const pending = this.pendingOpens.get(command.sessionId);
    if (pending && (command.kind === 'cancel' || command.kind === 'close')) {
      const matches = sameIdentity(pending.identity, command);
      if (matches) pending.canceled = true;
      post(matches
        ? {
            protocol: PULL_SESSION_PROTOCOL,
            kind: 'accepted',
            sessionId: command.sessionId,
            operationId: command.operationId,
            generation: command.generation,
            requestId: command.requestId,
            command: command.kind,
          }
        : this.staleLifecycleResponse(command));
      return Promise.resolve();
    }
    if (command.kind === 'cancel' || command.kind === 'close') {
      post({
        protocol: PULL_SESSION_PROTOCOL,
        kind: 'accepted',
        sessionId: command.sessionId,
        operationId: command.operationId,
        generation: command.generation,
        requestId: command.requestId,
        command: command.kind,
      });
      return Promise.resolve();
    }
    post({
      protocol: PULL_SESSION_PROTOCOL,
      kind: 'error',
      sessionId: command.sessionId,
      operationId: command.operationId,
      generation: command.generation,
      requestId: command.requestId,
      error: serializeWorkerError(new Error('slide pull session is not open')),
    });
    return Promise.resolve();
  }

  /** PullSessionHost has already reclaimed ownership when a data post throws. */
  async dispatchSafely(
    command: PullSessionCommand<number>,
    post: (response: PullSessionResponse<ArrayBuffer, number>, transfer?: Transferable[]) => void,
  ): Promise<void> {
    try {
      await this.dispatch(command, post);
    } catch (error) {
      try {
        post({
          protocol: PULL_SESSION_PROTOCOL,
          kind: 'error',
          sessionId: command.sessionId,
          operationId: command.operationId,
          generation: command.generation,
          requestId: command.requestId,
          error: serializeWorkerError(error),
        });
      } catch {
        // Host cleanup completed before this fallback was attempted.
      }
    }
  }

  /** Serialize sibling package work behind an open/unacknowledged slide. */
  run<T>(operation: () => T | Promise<T>): Promise<T> {
    if (this.lifecycleState !== 'ready') return Promise.reject(this.lifecycleError());
    const result = this.operationTail.then(() => this.coordinator.enqueue(async () => {
      if (this.resourceFailure) throw this.resourceFailure;
      return operation();
    }));
    const latched = result.catch((error: unknown) => {
      this.latchResourceFailure(error);
      throw error;
    });
    this.operationTail = latched.then(() => undefined, () => undefined);
    return latched;
  }

  /**
   * Close one complete presentation generation. The synchronous entrance
   * barrier prevents new work from escaping the snapshots below; concurrent
   * callers share the same convergence rather than starting another teardown.
   */
  reset(): Promise<void> {
    if (this.resetBarrier) return this.resetBarrier;
    this.lifecycleState = 'resetting';
    this.captureResetIdentities();
    const barrier = this.performReset().then(
      () => {
        this.resetIdentities.clear();
        this.lifecycleState = 'ready';
      },
      (error: unknown) => {
        this.lifecycleState = 'reset-failed';
        throw error;
      },
    ).finally(() => {
      if (this.resetBarrier === barrier) this.resetBarrier = undefined;
    });
    this.resetBarrier = barrier;
    return barrier;
  }

  private async performReset(): Promise<void> {
    for (const pending of this.pendingOpens.values()) pending.canceled = true;
    let requestId = 1;
    for (const { host, identity } of [...this.sessions.values()]) {
      let closeError: Error | undefined;
      await host.dispatch({
        protocol: PULL_SESSION_PROTOCOL,
        kind: 'close',
        ...identity,
        requestId: requestId++,
      }, (response) => {
        if (response.kind === 'error') closeError = deserializeWorkerError(response.error);
      });
      if (closeError) throw closeError;
    }
    this.sessions.clear();
    await this.operationTail;
    this.pendingOpens.clear();
    if (this.archive()) {
      await this.executeArchive((archive) => archive.close_presentation_session());
    }
    // Poison is scoped to one archive/presentation generation. Replacement is
    // safe only after every registered host and queued package operation above
    // has converged and the Rust generation has been closed.
    this.coordinatorGeneration = new PullSessionHostCoordinator();
    this.resourceFailure = undefined;
  }

  private assertReady(): void {
    if (this.lifecycleState !== 'ready') throw this.lifecycleError();
  }

  private lifecycleError(): Error {
    const failed = this.lifecycleState === 'reset-failed';
    const error = new Error(failed
      ? 'slide pull worker reset failed; retry reset before new work'
      : 'slide pull worker reset is in progress');
    error.name = 'PullSessionLifecycleError';
    return Object.assign(error, {
      code: failed ? RESET_FAILED_ERROR_CODE : RESETTING_ERROR_CODE,
    });
  }

  private captureResetIdentities(): void {
    for (const { identity } of this.sessions.values()) {
      this.resetIdentities.set(identity.sessionId, identity);
    }
    for (const { identity } of this.pendingOpens.values()) {
      this.resetIdentities.set(identity.sessionId, identity);
    }
  }

  private responseDuringReset(
    command: PullSessionCommand<number>,
  ): PullSessionResponse<ArrayBuffer, number> {
    if (command.kind === 'cancel' || command.kind === 'close') {
      const resetIdentity = this.resetIdentities.get(command.sessionId);
      if (resetIdentity && !sameIdentity(resetIdentity, command)) {
        return this.staleLifecycleResponse(command);
      }
      // Teardown owns cancellation of every captured operation. Correlated
      // lifecycle commands are idempotent observations, never forwarded to an
      // old host or to the archive generation being replaced.
      return {
        protocol: PULL_SESSION_PROTOCOL,
        kind: 'accepted',
        sessionId: command.sessionId,
        operationId: command.operationId,
        generation: command.generation,
        requestId: command.requestId,
        command: command.kind,
      };
    }
    return {
      protocol: PULL_SESSION_PROTOCOL,
      kind: 'error',
      sessionId: command.sessionId,
      operationId: command.operationId,
      generation: command.generation,
      requestId: command.requestId,
      error: serializeWorkerError(this.lifecycleError()),
    };
  }

  private requireArchive(): PptxSlideCursorArchive {
    const archive = this.archive();
    if (!archive) throw new Error('Presentation not loaded');
    return archive;
  }

  private async closeIdentity(identity: PullSessionIdentity<number>): Promise<void> {
    if (this.lifecycleState !== 'ready') return;
    const session = this.sessions.get(identity.sessionId);
    if (session) {
      await session.host.dispatch({
        protocol: PULL_SESSION_PROTOCOL,
        kind: 'close',
        ...identity,
        requestId: 1,
      }, () => undefined);
      return;
    }
    const pending = this.pendingOpens.get(identity.sessionId);
    if (pending && sameIdentity(pending.identity, identity)) pending.canceled = true;
  }

  private readResourceUsage(): OoxmlResourceUsageSnapshot | undefined {
    return readPptxSlideCursorUsage(this.executeArchive);
  }

  private latchResourceFailure(error: unknown): void {
    const typed = error instanceof OoxmlResourceLimitError ? error : parseResourceLimitError(error);
    if (typed) this.resourceFailure ??= typed;
  }

  private staleLifecycleResponse(
    command: Extract<PullSessionCommand<number>, { kind: 'cancel' | 'close' }>,
  ): PullSessionResponse<ArrayBuffer, number> {
    return {
      protocol: PULL_SESSION_PROTOCOL,
      kind: 'error',
      sessionId: command.sessionId,
      operationId: command.operationId,
      generation: command.generation,
      requestId: command.requestId,
      error: {
        message: 'stale lifecycle targets another slide operation',
        errorName: 'PullSessionProtocolError',
        code: 'ooxml-stale-lifecycle',
      },
    };
  }
}

export function isSlidePullCommand(value: unknown): value is PullSessionCommand<number> {
  return !!value && typeof value === 'object' &&
    (value as { protocol?: unknown }).protocol === PULL_SESSION_PROTOCOL;
}

function assertPositiveIdentity(identity: PullSessionIdentity<number>): void {
  if (!Number.isSafeInteger(identity.sessionId) || identity.sessionId <= 0) {
    throw new RangeError('session id must be a positive safe integer');
  }
  if (!Number.isSafeInteger(identity.operationId) || identity.operationId <= 0) {
    throw new RangeError('operation id must be a positive safe integer');
  }
  if (!Number.isSafeInteger(identity.generation) || identity.generation <= 0) {
    throw new RangeError('generation must be a positive safe integer');
  }
}

function sameIdentity(
  left: PullSessionIdentity<number>,
  right: Pick<PullSessionIdentity<number>, 'sessionId' | 'operationId' | 'generation'>,
): boolean {
  return left.sessionId === right.sessionId && left.operationId === right.operationId &&
    left.generation === right.generation;
}
