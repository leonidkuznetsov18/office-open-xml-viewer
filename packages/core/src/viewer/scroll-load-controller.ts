import type { DestroyableResource, TerminalResourceOwner } from '../internal/canvas-viewer-mechanics';

export interface ScrollLoadHooks<Resource extends DestroyableResource> {
  name(): string;
  borrowed(): boolean;
  borrowedMessage(): string;
  destroyed(): boolean;
  owner(): TerminalResourceOwner<Resource>;
  acquire(source: string | ArrayBuffer): Promise<Resource>;
  beforeReplace(previous: Resource | null): void;
  afterReplace(resource: Resource): void;
  mountOpeningWindow(): Promise<void>;
  selectionChanged(): void;
}

/** Atomic self-loaded resource replacement and opening-window lifecycle.
 * Failed acquisition retains the old resource; stale concurrent acquisitions
 * never commit or report through a replaced viewer. */
export class ScrollLoadController<Resource extends DestroyableResource> {
  constructor(private readonly hooks: ScrollLoadHooks<Resource>) {}

  async load(source: string | ArrayBuffer): Promise<void> {
    const name = this.hooks.name();
    if (this.hooks.destroyed()) throw new Error(`${name} is destroyed`);
    if (this.hooks.borrowed()) throw new Error(this.hooks.borrowedMessage());
    let selectionInvalidated = false;
    try {
      const resource = await this.hooks.owner().replace(
        () => this.hooks.acquire(source),
        (previous) => {
          this.hooks.beforeReplace(previous);
          selectionInvalidated = true;
        },
      );
      if (!resource) return;
      if (this.hooks.destroyed()) throw new Error(`${name} is destroyed`);
      this.hooks.afterReplace(resource);
      await this.hooks.mountOpeningWindow();
    } catch (error) {
      if (this.hooks.destroyed()) throw new Error(`${name} is destroyed`);
      throw error instanceof Error ? error : new Error(String(error));
    }
    // Consumer callbacks run only after the resource and first window commit;
    // their failures are not acquisition or render failures.
    if (selectionInvalidated && !this.hooks.destroyed()) this.hooks.selectionChanged();
  }
}
