import builtins
import clr
import collections
import fnmatch
import functools
import os
from pathlib import Path
# noinspection PyUnresolvedReferences
from typing import Optional, Sequence, TypingMeta, Optional, Iterable

# noinspection PyUnresolvedReferences
from System import Object
from dataclasses import dataclass

# noinspection PyUnreachableCode
if __debug__:
    clr.AddReference(
        os.path.abspath(os.path.join(__file__, "..", "../gen_lib/PyExcelLibrary/bin/Release", "PyExcelLibrary.dll")))
else:
    clr.AddReference(os.path.abspath(os.path.join(__file__, "..", "PyExcelLibrary.dll")))

# noinspection PyUnresolvedReferences
__all__ = ['Action', 'Actions', 'AddIn', 'AddIns', 'AddIns2', 'Adjustments', 'AllowEditRange', 'AllowEditRanges',
           'Application', 'Areas', 'Author', 'AutoCorrect', 'AutoFilter', 'AutoRecover', 'Border', 'Borders',
           'CalculatedFields', 'CalculatedItems', 'CalculatedMember', 'CalculatedMembers', 'CalloutFormat',
           'CellFormat', 'Characters', 'Chart', 'ChartArea', 'ChartColorFormat', 'ChartFillFormat', 'ChartFormat',
           'ChartGroup', 'ChartTitle', 'ColorFormat', 'Comment', 'CommentThreaded', 'Comments', 'CommentsThreaded',
           'Connections', 'ConnectorFormat', 'ControlFormat', 'Corners', 'CubeField', 'CubeFields', 'CustomProperties',
           'CustomProperty', 'CustomView', 'CustomViews', 'DataFeedConnection', 'DataTable',
           'DefaultPivotTableLayoutOptions', 'DefaultWebOptions', 'Diagram', 'DiagramNode', 'DiagramNodeChildren',
           'DiagramNodes', 'Dialog', 'DialogFrame', 'DialogSheet', 'Dialogs', 'DisplayFormat', 'DownBars', 'DropLines',
           'Error', 'ErrorCheckingOptions', 'Errors', 'FileExportConverter', 'FileExportConverters', 'FillFormat',
           'Filter', 'Filters', 'Floor', 'Font', 'FormatColor', 'FormatConditions', 'FreeformBuilder', 'Graphic',
           'GroupShapes', 'HPageBreak', 'HPageBreaks', 'HeaderFooter', 'HiLoLines', 'Hyperlink', 'Hyperlinks', 'Icon',
           'IconSet', 'IconSets', 'Interior', 'Legend', 'LineFormat', 'LinkFormat', 'ListColumn', 'ListColumns',
           'ListDataFormat', 'ListObject', 'ListObjects', 'ListRow', 'ListRows', 'Mailer', 'Menu', 'MenuBar',
           'MenuBars', 'MenuItem', 'MenuItems', 'Menus', 'Model', 'Model3DFormat', 'ModelConnection',
           'ModelFormatBoolean', 'ModelFormatCurrency', 'ModelFormatDate', 'ModelFormatDecimalNumber',
           'ModelFormatGeneral', 'ModelFormatPercentageNumber', 'ModelFormatScientificNumber', 'ModelFormatWholeNumber',
           'ModelMeasure', 'ModelMeasures', 'ModelRelationship', 'ModelRelationships', 'ModelTable', 'ModelTableColumn',
           'ModelTableColumns', 'ModelTables', 'Module', 'Modules', 'MultiThreadedCalculation', 'Name', 'Names',
           'ODBCConnection', 'ODBCError', 'ODBCErrors', 'OLEDBConnection', 'OLEDBError', 'OLEDBErrors', 'OLEFormat',
           'Outline', 'Page', 'PageSetup', 'Pages', 'Pane', 'Panes', 'Parameter', 'Parameters', 'Phonetic', 'Phonetics',
           'PictureFormat', 'PivotAxis', 'PivotCache', 'PivotCaches', 'PivotCell', 'PivotField', 'PivotFields',
           'PivotFilter', 'PivotFilters', 'PivotFormula', 'PivotFormulas', 'PivotItem', 'PivotItemList', 'PivotLayout',
           'PivotLine', 'PivotLineCells', 'PivotLines', 'PivotTable', 'PivotTableChangeList', 'PivotValueCell',
           'PlotArea', 'ProtectedViewWindow', 'ProtectedViewWindows', 'Protection', 'PublishObject', 'PublishObjects',
           'PublishedDoc', 'PublishedDocs', 'Queries', 'QueryTable', 'QueryTables', 'QuickAnalysis', 'RTD', 'Range',
           'Ranges', 'RecentFile', 'RecentFiles', 'Research', 'RoutingSlip', 'SeriesLines', 'ServerViewableItems',
           'ShadowFormat', 'Shape', 'ShapeNode', 'ShapeNodes', 'ShapeRange', 'Shapes', 'SheetViews', 'Sheets', 'Slicer',
           'SlicerCache', 'SlicerCacheLevel', 'SlicerCacheLevels', 'SlicerCaches', 'SlicerItem', 'SlicerItems',
           'SlicerPivotTables', 'Slicers', 'SmartTag', 'SmartTagAction', 'SmartTagActions', 'SmartTagOptions',
           'SmartTagRecognizer', 'SmartTagRecognizers', 'SmartTags', 'Sort', 'SortField', 'SortFields', 'SoundNote',
           'SparkAxes', 'SparkColor', 'SparkHorizontalAxis', 'SparkPoints', 'SparkVerticalAxis', 'Sparkline',
           'SparklineGroup', 'SparklineGroups', 'Speech', 'SpellingOptions', 'Style', 'Styles', 'Tab', 'TableObject',
           'TableStyle', 'TableStyleElement', 'TableStyleElements', 'TableStyles', 'TextConnection', 'TextEffectFormat',
           'TextFrame', 'TextFrame2', 'ThreeDFormat', 'TickLabels', 'TimelineState', 'TimelineViewState', 'Toolbar',
           'ToolbarButton', 'ToolbarButtons', 'Toolbars', 'TreeviewControl', 'UpBars', 'UsedObjects', 'UserAccess',
           'UserAccessList', 'VPageBreak', 'VPageBreaks', 'Validation', 'ValueChange', 'Walls', 'Watch', 'Watches',
           'WebOptions', 'Window', 'Windows', 'Workbook', 'WorkbookConnection', 'WorkbookQuery', 'Workbooks',
           'Worksheet', 'WorksheetDataConnection', 'WorksheetFunction', 'XPath', 'XmlDataBinding', 'XmlMap', 'XmlMaps',
           'XmlNamespace', 'XmlNamespaces', 'XmlSchema', 'XmlSchemas',
           'XlActionType', 'XlAllocation', 'XlAllocationMethod', 'XlAllocationValue', 'XlApplyNamesOrder',
           'XlArabicModes', 'XlArrangeStyle', 'XlAutoFillType', 'XlAutoFilterOperator', 'XlAxisGroup', 'XlBarShape',
           'XlBinsType', 'XlBorderWeight', 'XlBordersIndex', 'XlBuiltInDialog', 'XlCalcMemNumberFormatType',
           'XlCalculatedMemberType', 'XlCalculation', 'XlCalculationInterruptKey', 'XlCalculationState',
           'XlCategoryLabelLevel', 'XlCellChangedState', 'XlCellInsertionMode', 'XlCellType',
           'XlChartElementPosition', 'XlChartLocation', 'XlChartSplitType', 'XlChartType', 'XlCmdType',
           'XlColorIndex', 'XlCommandUnderlines', 'XlCommentDisplayMode', 'XlConnectionType',
           'XlConsolidationFunction', 'XlCopyPictureFormat', 'XlCreator', 'XlCredentialsMethod',
           'XlCubeFieldSubType', 'XlCubeFieldType', 'XlCutCopyMode', 'XlDVType', 'XlDataLabelsType',
           'XlDataSeriesDate', 'XlDataSeriesType', 'XlDirection', 'XlDisplayBlanksAs', 'XlDisplayDrawingObjects',
           'XlEditionOptionsOption', 'XlEditionType', 'XlEnableCancelKey', 'XlEnableSelection', 'XlFileAccess',
           'XlFileFormat', 'XlFileValidationPivotMode', 'XlFillWith', 'XlFilterAction', 'XlFilterStatus',
           'XlFixedFormatType', 'XlFormControl', 'XlFormatConditionType', 'XlFormulaLabel', 'XlGenerateTableRefs',
           'XlHAlign', 'XlHebrewModes', 'XlHtmlType', 'XlIconSet', 'XlLayoutFormType', 'XlLayoutRowType',
           'XlLegendPosition', 'XlLinkInfo', 'XlLinkType', 'XlListConflict', 'XlListDataType',
           'XlListObjectSourceType', 'XlLocationInTable', 'XlMSApplication', 'XlMailSystem', 'XlMousePointer',
           'XlOartHorizontalOverflow', 'XlOartVerticalOverflow', 'XlObjectSize', 'XlOrder', 'XlOrientation',
           'XlPTSelectionMode', 'XlPageBreak', 'XlPageBreakExtent', 'XlPageOrientation', 'XlPaperSize',
           'XlParameterDataType', 'XlParameterType', 'XlPasteSpecialOperation', 'XlPasteType', 'XlPictureAppearance',
           'XlPivotCellType', 'XlPivotFieldCalculation', 'XlPivotFieldDataType', 'XlPivotFieldOrientation',
           'XlPivotFieldRepeatLabels', 'XlPivotFilterType', 'XlPivotFormatType', 'XlPivotLineType',
           'XlPivotTableMissingItems', 'XlPivotTableSourceType', 'XlPivotTableVersionList', 'XlPlacement',
           'XlPlatform', 'XlPortugueseReform', 'XlPrintErrors', 'XlPrintLocation', 'XlPriority',
           'XlProtectedViewWindowState', 'XlPublishToDocsDisclosureScope', 'XlQueryType', 'XlQuickAnalysisMode',
           'XlRangeAutoFormat', 'XlReferenceStyle', 'XlRemoveDocInfoType', 'XlRobustConnect',
           'XlRoutingSlipDelivery', 'XlRoutingSlipStatus', 'XlRowCol', 'XlRunAutoMacro', 'XlSaveAsAccessMode',
           'XlSaveConflictResolution', 'XlSearchDirection', 'XlSeriesNameLevel', 'XlSheetType', 'XlSheetVisibility',
           'XlSizeRepresents', 'XlSlicerCacheType', 'XlSlicerCrossFilterType', 'XlSlicerSort',
           'XlSmartTagControlType', 'XlSmartTagDisplayMode', 'XlSortDataOption', 'XlSortMethod', 'XlSortOn',
           'XlSortOrder', 'XlSortOrientation', 'XlSourceType', 'XlSpanishModes', 'XlSparkScale', 'XlSparkType',
           'XlSparklineRowCol', 'XlSpeakDirection', 'XlSubscribeToFormat', 'XlSubtototalLocationType',
           'XlSummaryColumn', 'XlSummaryRow', 'XlTableStyleElementType', 'XlTextParsingType', 'XlTextQualifier',
           'XlTextVisualLayoutType', 'XlThemeColor', 'XlThemeFont', 'XlThreadMode', 'XlTickLabelOrientation',
           'XlTimelineLevel', 'XlToolbarProtection', 'XlTotalsCalculation', 'XlUpdateLinks', 'XlVAlign',
           'XlWebFormatting', 'XlWebSelectionType', 'XlWindowState', 'XlWindowType', 'XlWindowView',
           'XlXLMMacroType', 'XlXmlExportResult', 'XlXmlImportResult', 'XlYesNoGuess'
           ]

__all__ += ["Apps"]

# noinspection PyUnresolvedReferences
from PyExcelLibrary import (
    Action, Actions, AddIn, AddIns, AddIns2, Adjustments, AllowEditRange, AllowEditRanges, Areas, Author,
    AutoCorrect, AutoFilter, AutoRecover, Border, Borders, CalculatedFields, CalculatedItems, CalculatedMember,
    CalculatedMembers, CalloutFormat, CellFormat, Characters, Chart, ChartArea, ChartColorFormat, ChartFillFormat,
    ChartFormat, ChartGroup, ChartTitle, ColorFormat, Comment, CommentThreaded, Comments, CommentsThreaded, Connections,
    ConnectorFormat, ControlFormat, Corners, CubeField, CubeFields, CustomProperties, CustomProperty, CustomView,
    CustomViews, DataFeedConnection, DataTable, DefaultPivotTableLayoutOptions, DefaultWebOptions, Diagram, DiagramNode,
    DiagramNodeChildren, DiagramNodes, Dialog, DialogFrame, DialogSheet, Dialogs, DisplayFormat, DownBars, DropLines,
    Error, ErrorCheckingOptions, Errors, FileExportConverter, FileExportConverters, FillFormat, Filter, Filters, Floor,
    Font, FormatColor, FormatConditions, FreeformBuilder, Graphic, GroupShapes, HPageBreak, HPageBreaks, HeaderFooter,
    HiLoLines, Hyperlink, Hyperlinks, Icon, IconSet, IconSets, Interior, Legend, LineFormat, LinkFormat, ListColumn,
    ListColumns, ListDataFormat, ListObject, ListObjects, ListRow, ListRows, Mailer, Menu, MenuBar, MenuBars, MenuItem,
    MenuItems, Menus, Model, Model3DFormat, ModelConnection, ModelFormatBoolean, ModelFormatCurrency, ModelFormatDate,
    ModelFormatDecimalNumber, ModelFormatGeneral, ModelFormatPercentageNumber, ModelFormatScientificNumber,
    ModelFormatWholeNumber, ModelMeasure, ModelMeasures, ModelRelationship, ModelRelationships, ModelTable,
    ModelTableColumn, ModelTableColumns, ModelTables, Module, Modules, MultiThreadedCalculation, Name, Names,
    ODBCConnection, ODBCError, ODBCErrors, OLEDBConnection, OLEDBError, OLEDBErrors, OLEFormat, Outline, Page,
    PageSetup, Pages, Pane, Panes, Parameter, Parameters, Phonetic, Phonetics, PictureFormat, PivotAxis, PivotCache,
    PivotCaches, PivotCell, PivotField, PivotFields, PivotFilter, PivotFilters, PivotFormula, PivotFormulas, PivotItem,
    PivotItemList, PivotLayout, PivotLine, PivotLineCells, PivotLines, PivotTable, PivotTableChangeList, PivotValueCell,
    PlotArea, ProtectedViewWindow, ProtectedViewWindows, Protection, PublishObject, PublishObjects, PublishedDoc,
    PublishedDocs, Queries, QueryTable, QueryTables, QuickAnalysis, RTD, Ranges, RecentFile, RecentFiles,
    Research, RoutingSlip, SeriesLines, ServerViewableItems, ShadowFormat, Shape, ShapeNode, ShapeNodes, ShapeRange,
    Shapes, SheetViews, Slicer, SlicerCache, SlicerCacheLevel, SlicerCacheLevels, SlicerCaches, SlicerItem,
    SlicerItems, SlicerPivotTables, Slicers, SmartTag, SmartTagAction, SmartTagActions, SmartTagOptions,
    SmartTagRecognizer, SmartTagRecognizers, SmartTags, Sort, SortField, SortFields, SoundNote, SparkAxes, SparkColor,
    SparkHorizontalAxis, SparkPoints, SparkVerticalAxis, Sparkline, SparklineGroup, SparklineGroups, Speech,
    SpellingOptions, Style, Styles, Tab, TableObject, TableStyle, TableStyleElement, TableStyleElements, TableStyles,
    TextConnection, TextEffectFormat, TextFrame, TextFrame2, ThreeDFormat, TickLabels, TimelineState, TimelineViewState,
    Toolbar, ToolbarButton, ToolbarButtons, Toolbars, TreeviewControl, UpBars, UsedObjects, UserAccess, UserAccessList,
    VPageBreak, VPageBreaks, Validation, ValueChange, Walls, Watch, Watches, WebOptions, Window, Windows,
    WorkbookConnection, WorkbookQuery, WorksheetDataConnection, XPath,
    XmlDataBinding, XmlMap, XmlMaps, XmlNamespace, XmlNamespaces, XmlSchema, XmlSchemas,
)

from .enums import (
    XlActionType, XlAllocation, XlAllocationMethod, XlAllocationValue, XlApplyNamesOrder,
    XlArabicModes, XlArrangeStyle, XlAutoFillType, XlAutoFilterOperator, XlAxisGroup, XlBarShape,
    XlBinsType, XlBorderWeight, XlBordersIndex, XlBuiltInDialog, XlCalcMemNumberFormatType,
    XlCalculatedMemberType, XlCalculation, XlCalculationInterruptKey, XlCalculationState,
    XlCategoryLabelLevel, XlCellChangedState, XlCellInsertionMode, XlCellType,
    XlChartElementPosition, XlChartLocation, XlChartSplitType, XlChartType, XlCmdType,
    XlColorIndex, XlCommandUnderlines, XlCommentDisplayMode, XlConnectionType,
    XlConsolidationFunction, XlCopyPictureFormat, XlCreator, XlCredentialsMethod,
    XlCubeFieldSubType, XlCubeFieldType, XlCutCopyMode, XlDVType, XlDataLabelsType,
    XlDataSeriesDate, XlDataSeriesType, XlDirection, XlDisplayBlanksAs, XlDisplayDrawingObjects,
    XlEditionOptionsOption, XlEditionType, XlEnableCancelKey, XlEnableSelection, XlFileAccess,
    XlFileFormat, XlFileValidationPivotMode, XlFillWith, XlFilterAction, XlFilterStatus,
    XlFixedFormatType, XlFormControl, XlFormatConditionType, XlFormulaLabel, XlGenerateTableRefs,
    XlHAlign, XlHebrewModes, XlHtmlType, XlIconSet, XlLayoutFormType, XlLayoutRowType,
    XlLegendPosition, XlLinkInfo, XlLinkType, XlListConflict, XlListDataType,
    XlListObjectSourceType, XlLocationInTable, XlMSApplication, XlMailSystem, XlMousePointer,
    XlOartHorizontalOverflow, XlOartVerticalOverflow, XlObjectSize, XlOrder, XlOrientation,
    XlPTSelectionMode, XlPageBreak, XlPageBreakExtent, XlPageOrientation, XlPaperSize,
    XlParameterDataType, XlParameterType, XlPasteSpecialOperation, XlPasteType, XlPictureAppearance,
    XlPivotCellType, XlPivotFieldCalculation, XlPivotFieldDataType, XlPivotFieldOrientation,
    XlPivotFieldRepeatLabels, XlPivotFilterType, XlPivotFormatType, XlPivotLineType,
    XlPivotTableMissingItems, XlPivotTableSourceType, XlPivotTableVersionList, XlPlacement,
    XlPlatform, XlPortugueseReform, XlPrintErrors, XlPrintLocation, XlPriority,
    XlProtectedViewWindowState, XlPublishToDocsDisclosureScope, XlQueryType, XlQuickAnalysisMode,
    XlRangeAutoFormat, XlReferenceStyle, XlRemoveDocInfoType, XlRobustConnect,
    XlRoutingSlipDelivery, XlRoutingSlipStatus, XlRowCol, XlRunAutoMacro, XlSaveAsAccessMode,
    XlSaveConflictResolution, XlSearchDirection, XlSeriesNameLevel, XlSheetType, XlSheetVisibility,
    XlSizeRepresents, XlSlicerCacheType, XlSlicerCrossFilterType, XlSlicerSort,
    XlSmartTagControlType, XlSmartTagDisplayMode, XlSortDataOption, XlSortMethod, XlSortOn,
    XlSortOrder, XlSortOrientation, XlSourceType, XlSpanishModes, XlSparkScale, XlSparkType,
    XlSparklineRowCol, XlSpeakDirection, XlSubscribeToFormat, XlSubtototalLocationType,
    XlSummaryColumn, XlSummaryRow, XlTableStyleElementType, XlTextParsingType, XlTextQualifier,
    XlTextVisualLayoutType, XlThemeColor, XlThemeFont, XlThreadMode, XlTickLabelOrientation,
    XlTimelineLevel, XlToolbarProtection, XlTotalsCalculation, XlUpdateLinks, XlVAlign,
    XlWebFormatting, XlWebSelectionType, XlWindowState, XlWindowType, XlWindowView,
    XlXLMMacroType, XlXmlExportResult, XlXmlImportResult, XlYesNoGuess,
)

# noinspection PyUnreachableCode
if False:
    # noinspection PyUnresolvedReferences
    from PyExcelLibrary import ApplicationType, WorkbooksType, WorkbookType, SheetsType, WorksheetType, \
        RangeType, WorksheetFunctionType


@functools.wraps(builtins.repr)
def repr(obj):
    if isinstance(obj, ExcelType.Object):
        return str(obj)
    else:
        return builtins.repr(obj)


class ExcelType:
    # noinspection PyUnresolvedReferences
    from PyExcelLibrary import ExcelType as ET

    # noinspection PyUnresolvedReferences
    from System import Object

    # noinspection PyUnresolvedReferences
    from PyExcelLibrary import ExcelObject

    TypeArray1D = type(ET.Array1D)
    TypeArray2D = type(ET.Array2D)

    @classmethod
    def py_to_excel(cls, arg):
        if arg is None:
            return arg
        elif isinstance(arg, ImplCapsule):
            # XXX because, they require the raw COMObject
            # noinspection PyProtectedMember
            arg = arg.impl._object_

        return arg

    @classmethod
    def excel_to_py(cls, ret):
        if ret is None:
            return ret
        elif isinstance(ret, cls.TypeArray2D):
            ret = ExcelTable(ret)
        elif isinstance(ret, cls.Object):
            # TODO: check type and convert (check IMPL_REGISTERD set)
            raise TypeError(f"can't convert the {ret!r}")

        return ret

    del ET

    @classmethod
    def py_to_excel_range(cls, range: "Range", value):
        if isinstance(value, list) and value and isinstance(value[0], list):
            # TODO: fill range
            pass

        return False


TABLE_ROW = 0
TABLE_COL = 1


@dataclass
class ExcelTableInfo:
    row: int
    col: int
    row_range: range
    col_range: range
    row_base: int
    col_base: int


class ExcelTable:
    _impl: ExcelType.TypeArray2D
    info: ExcelTableInfo

    def __init__(self, impl):
        self.impl = impl

    @property
    def impl(self) -> ExcelType.TypeArray2D:
        return self._impl

    @impl.setter
    def impl(self, value: ExcelType.Object):
        self._impl = value
        self.refresh(value)

    def refresh(self, impl):
        row_range = range(impl.GetLowerBound(TABLE_ROW), impl.GetUpperBound(TABLE_ROW) + 1)
        col_range = range(impl.GetLowerBound(TABLE_COL), impl.GetUpperBound(TABLE_COL) + 1)

        self.info = ExcelTableInfo(
            row=len(row_range),
            col=len(col_range),
            row_range=row_range,
            col_range=col_range,
            row_base=row_range.start,
            col_base=col_range.start,
        )

    def __getitem__(self, item):
        info = self.info

        if isinstance(item, int):
            return [self.impl.GetValue(info.row_base, col) for col in info.col_range]

        x, y = item
        row = x + info.row_base
        col = y + info.col_base

        if row not in info.row_range:
            raise KeyError(item)
        elif col not in info.col_range:
            raise KeyError(item)

        return self.impl.GetValue(row, col)

    @property
    def row(self):
        return self.impl.GetLongLength(TABLE_ROW)

    @property
    def col(self):
        return self.impl.GetLongLength(TABLE_COL)

    def __iter__(self):
        info = self.info
        for row in info.row_range:
            yield [ExcelType.excel_to_py(self.impl.GetValue(row, col)) for col in info.col_range]

    def __repr__(self):
        return f"{type(self).__name__}[{self.row}, {self.col}]"


AppFolder = Path("working")

XLS_FILENAME_SUFFIX = (".xls", ".xlsx", ".xlsm")

revsered_ctx = {}

IMPL_REGISTERD = {
    "Application",
    "Workbooks",
    "Workbook",
    "Sheets",
    "Worksheet",
    "Range",
    "WorksheetFunction",
}


def patch_1():
    import PyExcelLibrary
    ctx = globals()
    for key in IMPL_REGISTERD:
        assert key not in ctx, key
        revsered_ctx[key] = getattr(PyExcelLibrary, key)
        ctx[key + "Type"] = type(key + "Type", (), {})


patch_1()


class ImplCapsule():
    __slots__ = "impl", "parent",

    def __init__(self, impl, parent=None):
        if impl is None:
            raise ValueError(f"{type(self).__name__}.impl object missing")

        object.__setattr__(self, "impl", impl)
        object.__setattr__(self, "parent", parent)

    def __getattr__(self, key):
        return getattr(self.impl, key)

    def __setattr__(self, key, value):
        if hasattr(type(self), key):
            super().__setattr__(key, value)
        else:
            setattr(self.impl, key, value)

    def __repr__(self):
        return f"<{repr(self.impl)}>"

    def __dir__(self):
        self_names = list(super().__dir__())
        impl_names = dir(self.impl)
        for name in (set(self_names) & set(impl_names)):
            impl_names.remove(name)

        return self_names + impl_names


class SequenceCapsule(ImplCapsule):
    def wrap(self, obj):
        raise NotImplementedError

    def get(self, item, default=None):
        if isinstance(item, int):
            if 0 <= item < len(self):
                obj = self.impl.get(item + 1)
                return self.wrap(obj)
            else:
                return default
        elif isinstance(item, str):
            for i in range(1, len(self) + 1):
                obj = self.impl.get(i)
                if obj.name == item:
                    return self.wrap(obj)
            else:
                return default
        else:
            raise TypeError(f"invaild index (int or str) = {item!r}")

    def __contains__(self, item):
        SENTINAL = object()
        return self.get(item, default=SENTINAL) is not SENTINAL

    def __len__(self):
        return self.impl.count

    def __getitem__(self, item):
        SENTINAL = object()
        obj = self.get(item, default=SENTINAL)
        if obj is SENTINAL:
            raise KeyError(item)

        return obj

    def __iter__(self):
        for i in range(1, len(self) + 1):
            obj = self.impl.get(i)
            yield self.wrap(obj)


class ExcelInstanceNotFound(Exception):
    pass


class Application(ImplCapsule, ApplicationType):
    impl: ApplicationType

    __slots__ = ImplCapsule.__slots__

    @property
    def app(self) -> "Application":
        return self

    @property
    def books(self):
        return Workbooks(self.impl.workbooks, self)

    @property
    def workbooks(self):
        return self.books

    @property
    def selection(self):
        return Range(self.impl.selection)

    @property
    def worksheet_function(self):
        return WorksheetFunction(self.impl.worksheet_function, self)


class Apps():
    # noinspection PyUnresolvedReferences
    from PyExcelLibrary.ExcelExtensions import ExcelAppCollection

    __slots__ = "impl",

    def __init__(self):
        self.impl = Apps.ExcelAppCollection()

    @property
    def session_id(self):
        return self.impl.SessionID

    def __iter__(self) -> Iterable["Application"]:
        for proc in self.impl.GetProcesses():
            impl = self.impl.FromProcess(proc)
            if impl is None:
                raise ExcelInstanceNotFound

            yield Application(impl, parent=proc)

    def from_process_id(self, pid):
        impl = self.impl.FromProcessID(pid)
        if impl is None:
            raise ExcelInstanceNotFound

        return Application(impl)

    def from_hwnd(self, hwnd):
        impl = self.impl.FromMainWindowHandle(hwnd)
        if impl is None:
            raise ExcelInstanceNotFound

        return Application(impl)

    @property
    def active(self) -> "Application":
        impl = self.impl.PrimaryInstance
        if impl is None:
            raise ExcelInstanceNotFound

        return Application(impl)

    @property
    def top_most(self):
        impl = self.impl.TopMostInstance
        if impl is None:
            raise ExcelInstanceNotFound

        return Application(impl)


class AppCapsule(ImplCapsule):
    @property
    def app(self) -> "Application":
        if self.parent is not None:
            return self.parent.app
        else:
            return Application(self.impl.application)


class Workbooks(AppCapsule, SequenceCapsule, Sequence["Workbook"]):
    impl: WorkbooksType
    parent: Application

    __slots__ = ImplCapsule.__slots__

    def __getitem__(self, item) -> "Workbook":
        return super().__getitem__(item)

    def __iter__(self) -> Iterable["Workbook"]:
        return map(self.wrap, iter(self.impl))

    def wrap(self, obj):
        return Workbook(obj, self)

    def find(self, *, name=None, path=None, folder=AppFolder) -> Optional["Workbook"]:
        if name is not None:
            for workbook in self.impl:
                if workbook.name == name:
                    return Workbook(workbook, self)
            else:
                if not name.endswith(XLS_FILENAME_SUFFIX):
                    raise ValueError(f"name {name!r} should endswith {XLS_FILENAME_SUFFIX}")

                path = (folder / f"{name}").absolute()
                if path.exists():
                    return Workbook(self.impl.open(str(path)), self)
                else:
                    path.absolute().parent.mkdir(parents=True, exist_ok=True)
                    workbook = self.impl.add()
                    workbook.save_as(str(path.absolute()))
                    return Workbook(workbook, self)
        elif path is not None:
            if isinstance(path, Path):
                path = str(path)

            for workbook in self.impl:
                if os.path.abspath(workbook.full_name) == os.path.abspath(path):
                    return Workbook(workbook, self)
            else:
                return Workbook(self.impl.open(path), self)


class Workbook(AppCapsule, WorkbookType):
    impl: WorkbookType
    parent: Workbooks

    __slots__ = ImplCapsule.__slots__

    @property
    def sheets(self):
        return Sheets(self.impl.sheets, self)

    @property
    def worksheets(self):
        return self.sheets

    @property
    def active_sheet(self):
        return Worksheet(self.impl.active_sheet, self)


class Sheets(AppCapsule, SequenceCapsule, SheetsType, Sequence["Worksheet"]):
    impl: SheetsType
    parent: Workbook

    __slots__ = ImplCapsule.__slots__

    def __getitem__(self, item) -> "Worksheet":
        return super().__getitem__(item)

    def __iter__(self) -> Iterable["Worksheet"]:
        return map(self.wrap, iter(self.impl))

    def wrap(self, obj):
        return Worksheet(obj, self)

    @property
    def active(self):
        return self.parent.active_sheet

    def glob(self, pattern):
        for sheet in self:  # type: Worksheet
            if fnmatch.fnmatch(sheet.name, pattern):
                yield sheet


class Worksheet(AppCapsule, WorksheetType):
    impl: WorksheetType
    parent: Sheets

    __slots__ = ImplCapsule.__slots__

    @property
    def cells(self):
        return Range(self.impl.cells, self)

    @property
    def active(self):
        selection = self.app.selection
        if selection.worksheet.name == self.impl.name:
            return Range(selection.impl, self)

        return None

    def get_range(self, cell1: object, cell2: object = None):
        return Range(self.impl.get_range(cell1, cell2), self)


def addr_row(n):
    return str(n + 1)


def addr_col(n):
    A = ord('A')

    if 0 <= n < 26:
        return chr(A + n)
    elif 26 <= n < 702:
        x, y = divmod(n, 26)
        return chr(A + x - 1) + chr(A + y)
    elif 702 <= n < 16384:
        x, n = divmod(n, 26 ** 2)
        y, z = divmod(n, 26)
        return chr(A + x - 1) + chr(A + y - 1) + chr(A + z)


# noinspection PyRedeclaration
class Range(AppCapsule, RangeType):
    impl: RangeType
    parent: Worksheet

    ITER_THRESHOLD_ALL_CELL = 2 ** 20
    ITER_THRESHOLD_ROW_CELL = 2 ** 18

    # https://msdn.microsoft.com/ko-kr/library/microsoft.office.interop.excel.range.aspx

    __slots__ = ImplCapsule.__slots__

    def __getitem__(self, key) -> "Range":
        if isinstance(key, tuple):
            klen = len(key)
            if klen == 1:
                key = key[0], None
            elif klen != 2:
                raise Exception

            row, column = key
            if isinstance(row, int) and isinstance(column, int):
                return Range(self.impl.get(row + 1, column + 1), self.parent)
            else:
                # noinspection PyUnresolvedReferences
                return Range(self.impl.get_item0(row, column), self.parent)
        else:
            # noinspection PyUnresolvedReferences
            return Range(self.impl.get_item0_key(key), self.parent)

    def __setitem__(self, key, value: "Range"):
        x, y = key
        if isinstance(x, int) and isinstance(y, int):
            self.impl.set(x + 1, y + 1, value)
        else:
            raise ValueError(f"unknown index {key!r}")

    @property
    def entire_row(self):
        return Range(self.impl.entire_row, self.parent)

    @property
    def entire_column(self):
        return Range(self.impl.entire_column, self.parent)

    @property
    def rows_count(self):
        # noinspection PyTypeChecker
        return int(self.impl.rows.count)

    @property
    def columns_count(self):
        # noinspection PyTypeChecker
        return int(self.impl.columns.count)

    # noinspection PyMethodOverriding
    def __iter__(self) -> Iterable["Range"]:
        if self.rows_count * self.columns_count > self.ITER_THRESHOLD_ALL_CELL:
            raise Exception

        for row in range(self.rows_count):
            for column in range(self.columns_count):
                yield self[row, column]

    @property
    def start(self):
        return self[0, 0]

    @property
    def end(self):
        return self[self.rows_count - 1, self.columns_count - 1]

    @property
    def header(self):
        return self[0, :self.columns_count]

    class ProxyDict(collections.Mapping):
        def __init__(self, header, row):
            self.header = header
            self.row = row

        def __getitem__(self, item):
            try:
                idx = self.header.index(item)
                return self.row[idx,]
            except IndexError:
                raise KeyError(item)

        def __len__(self):
            return len(self.header)

        def __iter__(self):
            return iter(self.header)

    def as_dict(self, raw=False, cache=True):
        header = self.header.value[0]
        table = self.data
        if raw:
            dict_cls = self.ProxyDict if cache else dict
            for row in table.rows:
                yield dict_cls(header, row)
        else:
            for row_values in table.value:
                yield dict(zip(header, row_values))

    @property
    def data(self):
        return self[1:, :]

    @property
    def rows(self):
        for row in range(self.rows_count):
            yield self[row, :]

    @property
    def columns(self):
        for column in range(self.columns_count):
            yield self[:, column]

    def get_range(self, cell1: object, cell2: object = None):
        return Range(self.impl.get_range(cell1, cell2), self.parent)

    @property
    def value0(self):
        return self.impl.value

    @value0.setter
    def value0(self, value):
        self.impl.value = value

    @property
    def value(self):
        value = self.impl.value
        value = ExcelType.excel_to_py(value)
        return value

    @value.setter
    def value(self, value):
        if ExcelType.py_to_excel_range(self, value):
            return

        value = ExcelType.py_to_excel(value)
        self.impl.value = value

    @value.deleter
    def value(self):
        self.impl.value = None

    def get_end(self, direction: XlDirection) -> RangeType:
        return Range(self.impl.get_end(direction), self.parent)


# noinspection PyRedeclaration
class WorksheetFunction(AppCapsule, WorksheetFunctionType):
    impl: WorksheetFunctionType
    parent: Worksheet

    def __getattr__(self, item):
        if hasattr(self.impl, item):
            return ExcelFunction(getattr(self.impl, item), self, name=item)

        # TODO: full exception?
        raise AttributeError(item)


class ExcelFunction(AppCapsule):
    impl: ExcelType.Object
    parent: WorksheetFunction
    name: str

    def __init__(self, impl, parent, *, name):
        object.__setattr__(self, "name", name)
        super().__init__(impl, parent)

    def __call__(self, *args):
        args = tuple(map(ExcelType.py_to_excel, args))
        ret = self.impl(*args)
        py_ret = ExcelType.excel_to_py(ret)
        return py_ret


def patch_2():
    ctx = globals()
    for key, value in revsered_ctx.items():
        ctx[key].__annotations__["impl"] = value


patch_2()

apps = Apps()
