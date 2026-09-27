import type { CjkLang } from '@silurus/ooxml-core';
import type {
  BodyElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocRun,
  ImageRun,
  ChartRun,
  ShapeRun,
  SectionProps,
} from '../types';
import type { ResolvedFontMetric } from '@silurus/ooxml-core';
import { type FloatRect, FLOAT_OVERLAP_EPS, isWrapFloat } from '../float-layout.js';
import {
  type FrameBox,
  computeFrameBox,
  frameXContainer,
  pushFloatRect,
  registerFrameFloat,
} from '../frame-geometry.js';
import { resolveFloatingTableBoxPt } from '../float-table-geometry.js';
import { xContainer, yContainer, resolveAnchorX, resolveAnchorY } from '../anchor-geometry.js';
import {
  resolveParagraphLayoutContext,
  resolveSectionLayoutContext,
  type DocumentLayoutSettings,
  type SectionLayoutContext,
} from '../layout-context.js';
import type {
  BlockLayoutAlgorithms,
  BodyFlowRegistryDeltaPt,
  BodyFlowRegistrySnapshotPt,
  DeepReadonly,
  DrawingMLCollisionRegistrySnapshotPt,
  LayoutServices,
  FloatRegistryEntryPt,
  FloatRegistrySnapshotPt,
  FloatingTablePlacementLayout,
  DrawingMLCollisionEntryPt,
  NoteLayout,
  ParagraphLayout,
  SourceRef,
  StoryBlockInput,
  StoryLayout,
  TableLayout,
  TableLayoutInput,
} from './types.js';
import {
  beginFloatingTablePlacementTransaction,
  floatingTableRegistryDelta,
  resolveFloatingTablePlacementInTransaction,
  validateFloatingTableRegistryDelta,
} from './floating-table-transaction.js';
import {
  floatRegistryParticipant,
  resolveBlockFlowAdmission,
  resolvePageAnchoredTableDeferral,
} from './floats.js';
import { ExactConvergenceError, convergeExactState } from './convergence.js';
import { LayoutInvariantError } from './diagnostics.js';
import type { LayoutOptions } from './options.js';
import {
  createLayoutServicesRuntimeView,
  fieldAcquisitionContextOf,
  verticalGlyphMeasurementServiceOf,
} from './runtime-state.js';
import { attachStoryBlockLayoutAlgorithms, layoutStory as layoutSharedStory } from './stories.js';
import {
  buildNoteNumberMap,
  footnoteIdsInRetainedLines,
  footnoteIdsInRetainedSlice,
  indexNotes,
  noteReferenceIdsInDocumentOrder,
} from './note-reference-ownership.js';
import type {
  BodyAcquisitionLocation,
  BodyLayoutKernel,
  BodyLayoutSession,
  PageAnchorPrescanInput,
  BodyParagraphAcquisitionInput,
  BodyTableAcquisitionInput,
} from './body-layout-kernel.js';
import { NoteCapacityExceededError } from './body-layout-kernel.js';
import { FlowCapacityExceededError } from './flow.js';
import { projectBodyOccurrence } from './occurrence-projection.js';
import { sectionBodyInsetPt as bodyMarginInsetPt, physicalSectionGeometry } from './context.js';
import {
  isAllRotatedVerticalTextDirection,
  isVerticalSection,
  isVerticalTextDirection,
  physicalLayoutSection,
  verticalLayoutSection,
} from './section-orientation.js';
import {
  gridForParagraphContext,
  paragraphMeasurementEnvironment,
} from './measurement-environment.js';
import { createRevisionAuthorColorResolver } from './track-changes.js';
import {
  BODY_STORY_CONTEXT,
  bodyAnchorReferenceFrames,
  retainedTableRecord,
  resolveBodyParagraphLayoutContext,
  resolveStateParagraphLayoutContext,
  withTableCellStory,
} from './acquisition-state.js';
import { applyNumberingBodyOffset, resolveNumberingMarkerGeometry } from './numbering-marker.js';
import { measureTableIntrinsicWidths, resolveTableColumnWidths } from './table-columns.js';
import {
  measureParagraphIntrinsicWidths,
  measureTableCellIntrinsicWidths,
} from './intrinsic-width.js';
import {
  buildFont,
  fontClassesWithPitches,
  getDefaultFontSize,
  paragraphMarkLineHeight,
} from '../line-layout.js';
import type { DocGridCtx } from '../line-layout.js';
import { measureParagraph } from '../paragraph-measure.js';
import {
  acquireRetainedTable,
  retainedTableAcquisitionIsReusableAcrossPages,
  type RetainedTableAcquisition,
} from './table-acquisition.js';
import { combineAdjacentTableLayoutInputs } from './adjacent-table-layout-input.js';
import { layoutTable as layoutRetainedTableInput } from './table.js';
import {
  startTableFragmentCursor,
  takeTableFragment,
  type PageDependentTableBlockRequest,
} from './table-pagination.js';
import { paragraphGapAdjustment } from './paragraph-spacing.js';
import {
  bottomBorderExtentPt,
  resolveParagraphBorderEdges,
  topBorderExtentPt,
  type ParagraphBorderEdges,
} from './paragraph-border-adjacency.js';
import {
  acquireParagraphResult,
  acquireRetainedFrameGroup,
  bodyFrameGroupFor,
  bodyParagraphBorderEdgesFor,
  projectPhysicalAnchorResult,
  retainedFrameMaximumBaselineLoweringPt,
  type BodyFrameGroup,
} from './paragraph.js';
import { wordLoweredDropCapAnchorLeadingPt } from './body-pagination-compatibility.js';
import type { CompleteTextBoxStoryAcquirer } from './paragraph.js';
import type {
  AnchorFloatRegistrationState,
  BodyAcquisitionState,
  BodyMeasurementContext,
  RetainedTableRecord,
} from './acquisition-context.js';
import {
  ownedParagraphAnchorCollisions,
  inheritedParagraphAuthorityForReacquisition,
  TRANSIENT_TABLE_FINAL_FRAME_EXCLUSION_PREFIX,
} from './paragraph-wrap-registry.js';
import { acquireRegisteredParagraph } from './registered-paragraph-acquisition.js';
import { paragraphAnchorCollisions, paragraphWrapExclusions } from './paragraph-float-authority.js';
import { bodyRootFloatingTablePlacementKey } from './source-key.js';
import {
  applyDrawingMLCollisionRegistryDelta,
  createDrawingMLCollisionRegistry,
  drawingMLCollisionRegistryDelta,
  validateDrawingMLCollisionRegistryDelta,
} from './drawingml-collision-registry.js';
import { resolveAnchorFrame } from './anchor-frame.js';
import { isPageLevelWrapFloat } from './anchor-classification.js';
import { physicalToLogicalAnchorBox } from '../vertical-text.js';
import type { MeasurementTextContext } from './measurement-capabilities.js';
import type {
  LayoutFlowBlock,
  LayoutParagraphBlock,
  LayoutSourceStore,
  LayoutStoryBlock,
  LayoutTableBlock,
} from './layout-source-store.js';
import type {
  ParagraphChartRun,
  ParagraphImageRun,
  ParagraphLayoutSource,
  ParagraphShapeRun,
} from './text.js';
import type { TableLayoutSource } from './table-source-acquisition.js';
import { collectBodyFrameGroups, prepareBodyFrameMetadata } from './frame.js';
import {
  physicalToLogicalMatrix,
  uprightPhysicalExtent,
  writingModeFromTextDirection,
} from './coordinate-space.js';

export interface BodyTableMeasurementContext {
  readonly dependencies: { readonly source: LayoutSourceStore };
  readonly services: LayoutServices;
  readonly state: BodyAcquisitionState;
  readonly sessionState: {
    location: BodyAcquisitionLocation;
    floatRegistry: FloatRegistrySnapshotPt;
    drawingCollisionRegistry: DrawingMLCollisionRegistrySnapshotPt;
  };
  readonly effectiveTablePositioning: LayoutSourceStore['acquisition']['effectiveTablePositioning'];
}

export interface BodyTableMeasurementOperations {
  readonly setBodyAcquisitionLocation: (
    services: LayoutServices,
    state: BodyAcquisitionState,
    sessionState: BodyTableMeasurementContext['sessionState'],
    next: BodyAcquisitionLocation,
  ) => void;
  readonly sourceElement: (store: LayoutSourceStore, ref: SourceRef) => LayoutFlowBlock;
  readonly computeTablePtLayout: (
    state: BodyAcquisitionState,
    table: TableLayoutSource,
    contentWPt: number,
    sourceIndex: number,
  ) => { colWidthsPt: number[]; rowContentHeightsPt: number[]; rowHeightsPt: number[] };
  readonly ordinaryAcquisitionInputForAdjacentGroup: (
    group: ReturnType<typeof combineAdjacentTableLayoutInputs>,
  ) => TableLayoutInput;
  readonly reacquireBodyTableBlock: (
    state: BodyAcquisitionState,
    store: LayoutSourceStore,
    request: PageDependentTableBlockRequest,
  ) => ParagraphLayout | TableLayout;
}

export function measureBodyTableEntry(
  context: BodyTableMeasurementContext,
  request: Parameters<NonNullable<BodyLayoutSession['measureTable']>>[0],
  operations: BodyTableMeasurementOperations,
): ReturnType<NonNullable<BodyLayoutSession['measureTable']>> {
  const { dependencies, services, state, sessionState, effectiveTablePositioning } = context;
  operations.setBodyAcquisitionLocation(services, state, sessionState, request.location);
  if (request.input.kind === 'adjacent-table-group') {
    return measureAdjacentTableGroup(context, request as AdjacentBodyTableRequest, operations);
  }
  const table = operations.sourceElement(dependencies.source, request.input.source);
  if (table.type !== 'table') throw new Error('Table source kind mismatch');
  const sourceIndex = request.input.source.path[0]!;
  operations.computeTablePtLayout(state, table, request.availableInlineExtentPt, sourceIndex);
  const retained = retainedTableRecord(state, sourceIndex).acquisition;
  if (request.cursor && request.cursor.kind !== 'table') {
    throw new Error('Ordinary table acquisition received an adjacent-group cursor');
  }
  const cursor = request.cursor?.cursor ?? startTableFragmentCursor();
  const pageHeightPt = state.pageH;
  const authoredPositioning = state.acquisitionInputs.tableFormatInput(table).positioning;
  if (authoredPositioning) {
    return measurePositionedTable(context, request as OrdinaryBodyTableRequest, operations, {
      table,
      sourceIndex,
      retained,
      cursor,
      pageHeightPt,
      authoredPositioning,
    });
  }
  if (state.verticalPhys && !effectiveTablePositioning(table)) {
    if (request.cursor) {
      throw new Error('An upright physical table must remain atomic');
    }
    const physical = state.verticalPhys;
    const tableWidthPt = retained.layout.columnWidthsPt.reduce((sum, width) => sum + width, 0);
    if (
      tableWidthPt > request.availableBlockExtentPt &&
      request.availableBlockExtentPt < request.freshPageBlockExtentPt
    ) {
      return Object.freeze({
        layout: retained.layout,
        blockExtentPt: 0,
        nextCursor: Object.freeze({ kind: 'table' as const, cursor }),
        requiresFreshFlowRegion: true,
      });
    }
    const physicalLeftPt =
      physical.physicalPageWidthPt - request.location.cursorPt.yPt - tableWidthPt;
    const physicalTopPt = request.location.cursorPt.xPt;
    const physicalBandHeightPt = Math.max(
      retained.layout.advancePt,
      physical.pageHeight - physical.marginTop - physical.marginBottom,
    );
    const flowDomainId = `upright-physical-page:${request.location.pageIndex}`;
    const upright = takeTableFragment(retained, startTableFragmentCursor(), {
      availableHeightPt: physicalBandHeightPt,
      freshPageHeightPt: physicalBandHeightPt,
      placement: {
        container: {
          id: flowDomainId,
          kind: 'body',
          bounds: { xPt: 0, yPt: 0, widthPt: tableWidthPt, heightPt: physicalBandHeightPt },
        },
        cursor: { xPt: 0, yPt: 0 },
        availableBounds: {
          xPt: 0,
          yPt: 0,
          widthPt: tableWidthPt,
          heightPt: physicalBandHeightPt,
        },
      },
      services,
      compatibility: 'word',
      oversizedRowPolicy: 'atomic',
      page: {
        physicalPageIndex: request.location.pageIndex,
        displayPageNumber: state.displayPageNumber ?? request.location.pageIndex + 1,
        occurrenceId: `${retained.input.id}:upright-page:${request.location.pageIndex}`,
      },
      floatingTableFrames: {
        page: { xPt: 0, yPt: 0, widthPt: physical.pageWidth, heightPt: physical.pageHeight },
        margin: {
          xPt: physical.marginLeft,
          yPt: physical.marginTop,
          widthPt: Math.max(0, physical.pageWidth - physical.marginLeft - physical.marginRight),
          heightPt: Math.max(0, physical.pageHeight - physical.marginTop - physical.marginBottom),
        },
        column: {
          xPt: physical.marginLeft,
          yPt: physical.marginTop,
          widthPt: Math.max(0, physical.pageWidth - physical.marginLeft - physical.marginRight),
          heightPt: Math.max(0, physical.pageHeight - physical.marginTop - physical.marginBottom),
        },
      },
      floatingTableRegistry: Object.freeze({
        coordinateSpace: 'upright-physical-page-points' as const,
        flowDomainId,
        entries: Object.freeze([]),
        nextParagraphId: 0,
      }),
      finalPlacementTranslationPt: { xPt: physicalLeftPt, yPt: physicalTopPt },
      reacquirePageDependentBlock: (request) =>
        operations.reacquireBodyTableBlock(state, dependencies.source, request),
    });
    if (!upright.fragment || upright.nextCursor || upright.requiresFreshPage) {
      throw new Error('Upright table final-frame layout must remain atomic');
    }
    return Object.freeze({
      layout: upright.fragment,
      blockExtentPt: tableWidthPt,
      nextCursor: null,
      placement: Object.freeze({
        coordinateSpace: 'upright-physical' as const,
        xPt: physicalLeftPt + upright.fragment.flowBounds.xPt,
        yPt: physicalTopPt + upright.fragment.flowBounds.yPt,
        sectionFlowOwnership: 'host-flow' as const,
      }),
    });
  }
  const result = takeTableFragment(retained, cursor, {
    availableHeightPt: request.availableBlockExtentPt,
    freshPageHeightPt: request.freshPageBlockExtentPt,
    placement: {
      container: {
        id: request.location.flowDomainId,
        kind: 'body',
        bounds: {
          xPt: 0,
          yPt: 0,
          widthPt: request.availableInlineExtentPt,
          heightPt: request.availableBlockExtentPt,
        },
      },
      cursor: { xPt: 0, yPt: 0 },
      availableBounds: {
        xPt: 0,
        yPt: 0,
        widthPt: request.availableInlineExtentPt,
        heightPt: request.availableBlockExtentPt,
      },
    },
    services,
    compatibility: 'word',
    page: {
      physicalPageIndex: request.location.pageIndex,
      displayPageNumber: request.location.pageIndex + 1,
      occurrenceId: `${retained.input.id}:body:${request.location.pageIndex}`,
    },
    floatingTableFrames: {
      page: { xPt: 0, yPt: 0, widthPt: state.pageWidth, heightPt: pageHeightPt },
      margin: {
        xPt: state.marginLeft,
        yPt: state.marginTop,
        widthPt: Math.max(0, state.pageWidth - state.marginLeft - state.marginRight),
        heightPt: Math.max(0, pageHeightPt - state.marginTop - state.marginBottom),
      },
      column: request.location.availableBounds,
    },
    floatingTableRegistry: sessionState.floatRegistry,
    finalPlacementTranslationPt: {
      xPt: request.location.availableBounds.xPt,
      yPt: request.location.cursorPt.yPt,
    },
    reacquirePageDependentBlock: (request) =>
      operations.reacquireBodyTableBlock(state, dependencies.source, request),
  });
  const tableInlineStartPt = request.location.availableBounds.xPt + retained.layout.flowBounds.xPt;
  const tableInlineEndPt = tableInlineStartPt + retained.layout.flowBounds.widthPt;
  const remainingTableExtentPt = result.fragment?.advancePt ?? 0;
  const retryAtBlockStartPt = resolveBlockFlowAdmission({
    inlineStartPt: tableInlineStartPt,
    inlineEndPt: tableInlineEndPt,
    blockStartPt: request.location.cursorPt.yPt,
    blockExtentPt: remainingTableExtentPt,
    blockers: sessionState.floatRegistry.entries.map(floatRegistryParticipant),
    overlapEpsilonPt: FLOAT_OVERLAP_EPS,
  }).blockStartPt;
  if (retryAtBlockStartPt > request.location.cursorPt.yPt) {
    return Object.freeze({
      layout: retained.layout,
      blockExtentPt: 0,
      nextCursor: request.cursor ?? null,
      retryAtBlockStartPt,
    });
  }
  if (!result.fragment || result.requiresFreshPage) {
    return Object.freeze({
      layout: retained.layout,
      blockExtentPt: 0,
      nextCursor: Object.freeze({ kind: 'table' as const, cursor }),
      requiresFreshFlowRegion: true,
    });
  }
  return Object.freeze({
    layout: result.fragment,
    blockExtentPt: result.fragment.advancePt,
    ...(result.fragment.unpaintedOverflowPt !== undefined
      ? { unpaintedOverflowPt: result.fragment.unpaintedOverflowPt }
      : {}),
    nextCursor: result.nextCursor
      ? Object.freeze({ kind: 'table' as const, cursor: result.nextCursor })
      : null,
    ...(result.floatingTableRegistryDelta
      ? {
          flowRegistryDelta: Object.freeze({
            floats: result.floatingTableRegistryDelta,
          }),
        }
      : {}),
  });
}

type BodyTableRequest = Parameters<NonNullable<BodyLayoutSession['measureTable']>>[0];
type TableMeasureResult = ReturnType<NonNullable<BodyLayoutSession['measureTable']>>;
type AdjacentBodyTableRequest = BodyTableRequest & {
  input: Extract<BodyTableRequest['input'], { kind: 'adjacent-table-group' }>;
};
type OrdinaryBodyTableRequest = BodyTableRequest & {
  input: Exclude<BodyTableRequest['input'], { kind: 'adjacent-table-group' }>;
};
interface PositionedTableFrame {
  readonly table: TableLayoutSource;
  readonly sourceIndex: number;
  readonly retained: RetainedTableAcquisition;
  readonly cursor: ReturnType<typeof startTableFragmentCursor>;
  readonly pageHeightPt: number;
  readonly authoredPositioning: NonNullable<
    ReturnType<BodyAcquisitionState['acquisitionInputs']['tableFormatInput']>['positioning']
  >;
}

function measureAdjacentTableGroup(
  context: BodyTableMeasurementContext,
  request: AdjacentBodyTableRequest,
  operations: BodyTableMeasurementOperations,
): TableMeasureResult {
  const { dependencies, services, state } = context;

  if (request.cursor && request.cursor.kind !== 'adjacent-table-group') {
    throw new Error('Adjacent table group acquisition received an ordinary table cursor');
  }
  const records = request.input.tables.map((tableInput) => {
    const table = operations.sourceElement(dependencies.source, tableInput.source);
    if (table.type !== 'table') throw new Error('Table source kind mismatch');
    const sourceIndex = tableInput.source.path[0]!;
    operations.computeTablePtLayout(state, table, request.availableInlineExtentPt, sourceIndex);
    return retainedTableRecord(state, sourceIndex).acquisition;
  });
  const combinedInput = operations.ordinaryAcquisitionInputForAdjacentGroup(
    combineAdjacentTableLayoutInputs(
      request.input.logicalSequenceId,
      records.map((record) => record.input),
    ),
  );
  const placement = {
    container: {
      id: request.location.flowDomainId,
      kind: 'body' as const,
      bounds: {
        xPt: 0,
        yPt: 0,
        widthPt: request.availableInlineExtentPt,
        heightPt: request.freshPageBlockExtentPt,
      },
    },
    cursor: { xPt: 0, yPt: 0 },
    availableBounds: {
      xPt: 0,
      yPt: 0,
      widthPt: request.availableInlineExtentPt,
      heightPt: request.freshPageBlockExtentPt,
    },
  };
  const combinedLayout = layoutRetainedTableInput(combinedInput, placement, services).layout;
  const nestedById: Record<string, RetainedTableAcquisition> = {};
  records.forEach((record) =>
    Object.entries(record.nestedById).forEach(([id, nested]) => {
      if (nestedById[id] && nestedById[id] !== nested) {
        throw new Error(`Adjacent table group has duplicate nested table id: ${id}`);
      }
      nestedById[id] = nested;
    }),
  );
  const combined: RetainedTableAcquisition = Object.freeze({
    input: combinedInput,
    layout: combinedLayout,
    nestedById: Object.freeze(nestedById),
    floatingTables: Object.freeze(records.flatMap((record) => record.floatingTables)),
  });
  const groupCursor: import('./body-layout-kernel.js').AdjacentTableGroupCursor =
    request.cursor?.cursor ??
    Object.freeze({
      tableIndex: 0,
      sourceRowIndex: 0,
    });
  const rowsBefore = request.input.tables
    .slice(0, groupCursor.tableIndex)
    .reduce((sum, tableInput) => sum + (tableInput.rowCount ?? 0), 0);
  const globalRowIndex = rowsBefore + groupCursor.sourceRowIndex;
  const cursor =
    groupCursor.tableCursor ??
    Object.freeze({
      ...startTableFragmentCursor(),
      rowIndex: globalRowIndex,
    });
  if (cursor.rowIndex !== globalRowIndex) {
    throw new Error('Adjacent-table group and table-fragment cursors disagree');
  }
  const result = takeTableFragment(combined, cursor, {
    availableHeightPt: request.availableBlockExtentPt,
    freshPageHeightPt: request.freshPageBlockExtentPt,
    placement,
    services,
    compatibility: 'word',
    page: {
      physicalPageIndex: request.location.pageIndex,
      displayPageNumber: request.location.pageIndex + 1,
      occurrenceId: `${combinedInput.id}:body:${request.location.pageIndex}`,
    },
  });
  if (!result.fragment || result.requiresFreshPage) {
    return Object.freeze({
      layout: combined.layout,
      blockExtentPt: 0,
      nextCursor: Object.freeze({
        kind: 'adjacent-table-group' as const,
        cursor: groupCursor,
      }),
      requiresFreshFlowRegion: true,
    });
  }
  const nextGroupCursor = result.nextCursor
    ? (() => {
        let tableIndex = 0;
        let firstRow = 0;
        while (tableIndex < request.input.tables.length) {
          const rowCount = request.input.tables[tableIndex]!.rowCount ?? 0;
          if (result.nextCursor!.rowIndex < firstRow + rowCount) break;
          firstRow += rowCount;
          tableIndex += 1;
        }
        if (tableIndex >= request.input.tables.length) return null;
        return Object.freeze({
          tableIndex,
          sourceRowIndex: result.nextCursor!.rowIndex - firstRow,
          tableCursor: result.nextCursor!,
        });
      })()
    : null;
  return Object.freeze({
    layout: result.fragment,
    blockExtentPt: result.fragment.advancePt,
    ...(result.fragment.unpaintedOverflowPt !== undefined
      ? { unpaintedOverflowPt: result.fragment.unpaintedOverflowPt }
      : {}),
    nextCursor: nextGroupCursor
      ? Object.freeze({ kind: 'adjacent-table-group' as const, cursor: nextGroupCursor })
      : null,
    ...(result.floatingTableRegistryDelta
      ? {
          flowRegistryDelta: Object.freeze({
            floats: result.floatingTableRegistryDelta,
          }),
        }
      : {}),
  });
}

function measurePositionedTable(
  context: BodyTableMeasurementContext,
  request: OrdinaryBodyTableRequest,
  operations: BodyTableMeasurementOperations,
  frame: PositionedTableFrame,
): TableMeasureResult {
  const { dependencies, services, state, sessionState } = context;
  const { table, sourceIndex, retained, cursor, pageHeightPt, authoredPositioning } = frame;

  const positioning =
    request.cursor?.kind === 'table' && request.cursor.floatingContinuationFrame === 'fresh-text'
      ? Object.freeze({ ...authoredPositioning, vertAnchor: 'text', yPt: 0, yAlign: undefined })
      : authoredPositioning;
  const tableWidthPt = retained.layout.columnWidthsPt.reduce((sum, width) => sum + width, 0);
  const frames = Object.freeze({
    page: Object.freeze({ xPt: 0, yPt: 0, widthPt: state.pageWidth, heightPt: pageHeightPt }),
    margin: Object.freeze({
      xPt: state.marginLeft,
      yPt: state.marginTop,
      widthPt: Math.max(0, state.pageWidth - state.marginLeft - state.marginRight),
      heightPt: Math.max(0, pageHeightPt - state.marginTop - state.marginBottom),
    }),
    text: Object.freeze({
      xPt: request.location.cursorPt.xPt,
      yPt: request.location.cursorPt.yPt,
      widthPt: request.availableInlineExtentPt,
      heightPt: retained.layout.advancePt,
    }),
  });
  const raw = resolveFloatingTableBoxPt(
    positioning,
    frames,
    tableWidthPt,
    retained.layout.advancePt,
  );
  const ownPrescanOccurrenceId = bodyRootFloatingTablePlacementKey(
    request.input.source,
    request.location.pageIndex,
    cursor.rowIndex,
    cursor.rowFragmentIndex,
  );
  // The advance registration is for text before this table. Its
  // nested contents must acquire against other floats, not against
  // the table that owns them.
  const hasOwnPrescan =
    (positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin') &&
    sessionState.floatRegistry.entries.some(
      (entry) => entry.occurrenceId === ownPrescanOccurrenceId,
    );
  const nestedAcquisitionRegistry = hasOwnPrescan
    ? Object.freeze({
        ...sessionState.floatRegistry,
        entries: Object.freeze(
          sessionState.floatRegistry.entries.filter(
            (entry) => entry.occurrenceId !== ownPrescanOccurrenceId,
          ),
        ),
      })
    : sessionState.floatRegistry;
  const pageAnchoredCollision =
    request.cursor?.kind !== 'table' &&
    (positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin') &&
    resolvePageAnchoredTableDeferral({
      bounds: {
        xPt: raw.x,
        yPt: raw.y,
        widthPt: raw.w,
        heightPt: raw.h,
      },
      blockers: sessionState.floatRegistry.entries
        .filter((entry) => entry.occurrenceId !== ownPrescanOccurrenceId)
        .map(floatRegistryParticipant),
      overlapEpsilonPt: FLOAT_OVERLAP_EPS,
    }).defer;
  if (pageAnchoredCollision) {
    // `word-page-anchored-table-collision-deferral`: a fresh page
    // preserves the authored absolute anchor instead of converting
    // the colliding table to a text continuation.
    return Object.freeze({
      layout: retained.layout,
      blockExtentPt: 0,
      nextCursor: Object.freeze({
        kind: 'table' as const,
        cursor,
        floatingContinuationFrame: 'authored' as const,
      }),
      requiresFreshFlowRegion: true,
    });
  }
  const absoluteAnchorMustSplit =
    (positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin') &&
    retained.layout.advancePt > request.freshPageBlockExtentPt;
  const admissionBlockEndPt = absoluteAnchorMustSplit
    ? request.location.availableBounds.yPt + request.location.availableBounds.heightPt
    : positioning.vertAnchor === 'page'
      ? frames.page.yPt + frames.page.heightPt
      : positioning.vertAnchor === 'margin'
        ? frames.margin.yPt + frames.margin.heightPt
        : request.location.availableBounds.yPt + request.location.availableBounds.heightPt;
  const freshAdmissionHeightPt = absoluteAnchorMustSplit
    ? request.freshPageBlockExtentPt
    : positioning.vertAnchor === 'page'
      ? frames.page.heightPt
      : positioning.vertAnchor === 'margin'
        ? frames.margin.heightPt
        : request.freshPageBlockExtentPt;
  const transaction = convergeFloatingParentTransaction(context, request, operations, {
    table,
    retained,
    cursor,
    positioning,
    frames,
    raw,
    nestedAcquisitionRegistry,
    admissionBlockEndPt,
    freshAdmissionHeightPt,
  });
  if (transaction.kind === 'fresh-flow-region') {
    return Object.freeze({
      layout: retained.layout,
      blockExtentPt: 0,
      nextCursor: Object.freeze({
        kind: 'table' as const,
        cursor,
        floatingContinuationFrame: 'fresh-text' as const,
      }),
      requiresFreshFlowRegion: true,
    });
  }
  const { result, fragment, resolved, nestedEntries } = transaction;
  const isFloatingContinuation =
    request.cursor?.kind === 'table' && request.cursor.floatingContinuationFrame !== undefined;
  const admittedBlockEndPt =
    request.location.availableBounds.yPt + request.location.availableBounds.heightPt;
  const hostFlowPlacements = [
    ...(fragment.resolvedFloatingTables ?? []),
    resolved.placement,
  ].filter((placement) => placement.source.positioning.vertAnchor === 'text');
  if (
    !isFloatingContinuation &&
    hostFlowPlacements.some(
      (placement) =>
        placement.exclusionBounds.yPt + placement.exclusionBounds.heightPt > admittedBlockEndPt,
    )
  ) {
    return Object.freeze({
      layout: fragment,
      blockExtentPt: 0,
      nextCursor: Object.freeze({
        kind: 'table' as const,
        cursor,
        floatingContinuationFrame: 'fresh-text' as const,
      }),
      requiresFreshFlowRegion: true,
    });
  }
  return Object.freeze({
    layout: fragment,
    blockExtentPt: 0,
    nextCursor: result.nextCursor
      ? Object.freeze({
          kind: 'table' as const,
          cursor: result.nextCursor,
          floatingContinuationFrame: 'fresh-text' as const,
        })
      : null,
    flowRegistryDelta: Object.freeze({
      floats: floatingTableRegistryDelta(
        sessionState.floatRegistry,
        Object.freeze([...nestedEntries, ...resolved.transaction.delta]),
        resolved.transaction.nextParagraphId,
      ),
    }),
    placement: Object.freeze({
      coordinateSpace: 'logical-body' as const,
      xPt: resolved.placement.xPt,
      yPt: resolved.placement.yPt,
      sectionFlowOwnership:
        positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin'
          ? ('page' as const)
          : ('host-flow' as const),
    }),
  });
}

type FloatingParentTransactionPass =
  | Readonly<{
      kind: 'fresh-flow-region';
      result: ReturnType<typeof takeTableFragment>;
    }>
  | Readonly<{
      kind: 'candidate';
      parentFrame: Readonly<{ xPt: number; yPt: number }>;
      result: ReturnType<typeof takeTableFragment>;
      fragment: NonNullable<ReturnType<typeof takeTableFragment>['fragment']>;
      resolved: ReturnType<typeof resolveFloatingTablePlacementInTransaction>;
      nestedEntries: readonly FloatRegistryEntryPt[];
      fingerprint: string;
    }>;
interface FloatingConvergenceFrame {
  readonly table: TableLayoutSource;
  readonly retained: RetainedTableAcquisition;
  readonly cursor: ReturnType<typeof startTableFragmentCursor>;
  readonly positioning: PositionedTableFrame['authoredPositioning'];
  readonly frames: Parameters<typeof resolveFloatingTableBoxPt>[1];
  readonly raw: ReturnType<typeof resolveFloatingTableBoxPt>;
  readonly nestedAcquisitionRegistry: FloatRegistrySnapshotPt;
  readonly admissionBlockEndPt: number;
  readonly freshAdmissionHeightPt: number;
}

function convergeFloatingParentTransaction(
  context: BodyTableMeasurementContext,
  request: OrdinaryBodyTableRequest,
  operations: BodyTableMeasurementOperations,
  frame: FloatingConvergenceFrame,
): FloatingParentTransactionPass {
  const { dependencies, services, state, sessionState } = context;
  const {
    table,
    retained,
    cursor,
    positioning,
    frames,
    raw,
    nestedAcquisitionRegistry,
    admissionBlockEndPt,
    freshAdmissionHeightPt,
  } = frame;
  let transaction: FloatingParentTransactionPass;
  try {
    transaction = convergeExactState<FloatingParentTransactionPass>({
      step: (previous) => {
        if (previous?.kind === 'fresh-flow-region') return previous;
        if (
          previous?.kind === 'candidate' &&
          previous.resolved.placement.xPt === previous.parentFrame.xPt &&
          previous.resolved.placement.yPt === previous.parentFrame.yPt
        ) {
          return previous;
        }
        const parentFrame = previous?.resolved.placement ?? {
          xPt: raw.x,
          yPt: raw.y,
        };
        const availableHeightPt = Math.max(0, admissionBlockEndPt - parentFrame.yPt);
        const result = takeTableFragment(retained, cursor, {
          availableHeightPt,
          freshPageHeightPt: freshAdmissionHeightPt,
          placement: {
            container: {
              id: `${request.location.flowDomainId}:floating-table`,
              kind: 'body',
              bounds: {
                xPt: 0,
                yPt: 0,
                widthPt: request.availableInlineExtentPt,
                heightPt: availableHeightPt,
              },
            },
            cursor: { xPt: 0, yPt: 0 },
            availableBounds: {
              xPt: 0,
              yPt: 0,
              widthPt: request.availableInlineExtentPt,
              heightPt: availableHeightPt,
            },
          },
          services,
          compatibility: 'word',
          oversizedRowPolicy: 'atomic',
          page: {
            physicalPageIndex: request.location.pageIndex,
            displayPageNumber: state.displayPageNumber ?? request.location.pageIndex + 1,
            occurrenceId: `${retained.input.id}:fitting-outer:${request.location.pageIndex}:${cursor.rowIndex}:${cursor.rowFragmentIndex}`,
          },
          floatingTableFrames: {
            page: frames.page,
            margin: frames.margin,
            column: frames.text,
          },
          floatingTableRegistry: nestedAcquisitionRegistry,
          finalPlacementTranslationPt: parentFrame,
          reacquirePageDependentBlock: (request) =>
            operations.reacquireBodyTableBlock(state, dependencies.source, request),
        });
        if (!result.fragment || result.requiresFreshPage) {
          return Object.freeze({
            kind: 'fresh-flow-region' as const,
            result,
          });
        }
        const sourcePlacement: FloatingTablePlacementLayout = Object.freeze({
          kind: 'floating-table-placement',
          occurrenceId: bodyRootFloatingTablePlacementKey(
            request.input.source,
            request.location.pageIndex,
            cursor.rowIndex,
            cursor.rowFragmentIndex,
          ),
          ownership: 'source',
          physicalPageIndex: request.location.pageIndex,
          displayPageNumber: state.displayPageNumber ?? request.location.pageIndex + 1,
          hostCellId: request.location.flowDomainId,
          sourceBlockIndex: request.input.source.path[0]!,
          anchorBlockIndex: request.input.source.path[0]!,
          tableId: result.fragment.id,
          overlap: table.overlap === 'never' ? 'never' : 'overlap',
          positioning,
          anchorBounds: frames.text,
          child: result.fragment,
        });
        const nestedEntries = result.floatingTableRegistryDelta?.entries ?? [];
        const nestedNextParagraphId =
          result.floatingTableRegistryDelta?.nextParagraphId ??
          sessionState.floatRegistry.nextParagraphId;
        const resolved = resolveFloatingTablePlacementInTransaction(
          sourcePlacement,
          frames,
          beginFloatingTablePlacementTransaction(
            sessionState.floatRegistry.entries,
            nestedNextParagraphId,
            sessionState.floatRegistry.coordinateSpace,
            sessionState.floatRegistry.flowDomainId,
          ),
        );
        const fingerprint = JSON.stringify({
          parentFrame: {
            xPt: resolved.placement.xPt,
            yPt: resolved.placement.yPt,
          },
          fragment: result.fragment,
          nestedEntries,
          resolvedBounds: resolved.placement.bounds,
        });
        return Object.freeze({
          kind: 'candidate' as const,
          parentFrame: Object.freeze({
            xPt: parentFrame.xPt,
            yPt: parentFrame.yPt,
          }),
          result,
          fragment: result.fragment,
          resolved,
          nestedEntries,
          fingerprint,
        });
      },
      stateOf: (value) =>
        value.kind === 'fresh-flow-region' ? 'fresh-flow-region' : value.fingerprint,
      limit: 16,
    }).value;
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        error.reason === 'cycle'
          ? 'Floating table parent/child transaction repeated an exact-state cycle'
          : 'Floating table parent/child transaction reached the operational pass limit 16',
      );
    }
    throw error;
  }
  return transaction;
}
