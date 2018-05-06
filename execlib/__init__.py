import builtins
import clr
import functools
import os
from pathlib import Path
# noinspection PyUnresolvedReferences
from typing import Optional, Sequence, TypingMeta

# noinspection PyUnresolvedReferences
from System import Object

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
           'XmlNamespace', 'XmlNamespaces', 'XmlSchema', 'XmlSchemas']

__all__ += ["Apps"]

# noinspection PyUnresolvedReferences
from PyExcelLibrary import (
    Action, Actions, AddIn, AddIns, AddIns2, Adjustments, AllowEditRange, AllowEditRanges, Application, Areas, Author,
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
    PublishedDocs, Queries, QueryTable, QueryTables, QuickAnalysis, RTD, Range, Ranges, RecentFile, RecentFiles,
    Research, RoutingSlip, SeriesLines, ServerViewableItems, ShadowFormat, Shape, ShapeNode, ShapeNodes, ShapeRange,
    Shapes, SheetViews, Sheets, Slicer, SlicerCache, SlicerCacheLevel, SlicerCacheLevels, SlicerCaches, SlicerItem,
    SlicerItems, SlicerPivotTables, Slicers, SmartTag, SmartTagAction, SmartTagActions, SmartTagOptions,
    SmartTagRecognizer, SmartTagRecognizers, SmartTags, Sort, SortField, SortFields, SoundNote, SparkAxes, SparkColor,
    SparkHorizontalAxis, SparkPoints, SparkVerticalAxis, Sparkline, SparklineGroup, SparklineGroups, Speech,
    SpellingOptions, Style, Styles, Tab, TableObject, TableStyle, TableStyleElement, TableStyleElements, TableStyles,
    TextConnection, TextEffectFormat, TextFrame, TextFrame2, ThreeDFormat, TickLabels, TimelineState, TimelineViewState,
    Toolbar, ToolbarButton, ToolbarButtons, Toolbars, TreeviewControl, UpBars, UsedObjects, UserAccess, UserAccessList,
    VPageBreak, VPageBreaks, Validation, ValueChange, Walls, Watch, Watches, WebOptions, Window, Windows, Workbook,
    WorkbookConnection, WorkbookQuery, Workbooks, Worksheet, WorksheetDataConnection, WorksheetFunction, XPath,
    XmlDataBinding, XmlMap, XmlMaps, XmlNamespace, XmlNamespaces, XmlSchema, XmlSchemas,
)


@functools.wraps(builtins.repr)
def repr(obj):
    if isinstance(obj, Object):
        return str(obj)
    else:
        return builtins.repr(obj)


AppFolder = Path("working")

XLS_FILENAME_SUFFIX = (".xls", ".xlsx", ".xlsm")

revsered_ctx = {}


def patch_1():
    target = {
        "Application",
        "Workbooks",
        "Workbook",
        "Sheets",
        "Worksheet",
        "Range",
    }

    ctx = globals()
    for key in target:
        revsered_ctx[key], ctx[key] = ctx[key], object


patch_1()


class ImplCapsule():
    __slots__ = "impl", "parent",

    def __init__(self, impl, parent=None):
        object.__setattr__(self, "impl", impl)
        object.__setattr__(self, "parent", parent)

    def __getattr__(self, key):
        return getattr(self.impl, key)

    def __setattr__(self, key, value):
        setattr(self.impl, key, value)

    def __delattr__(self, key):
        delattr(self.impl, key)

    def __repr__(self):
        return f"<{repr(self.impl)}>"

    def __iter__(self):
        return iter(self.impl)

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


class Apps:
    # noinspection PyUnresolvedReferences
    from PyExcelLibrary.ExcelExtensions import ExcelAppCollection

    __slots__ = "impl",

    def __init__(self):
        self.impl = Apps.ExcelAppCollection()

    @property
    def session_id(self):
        return self.impl.SessionID

    def __iter__(self):
        for proc in self.impl.GetProcesses():
            yield Application(self.impl.FromProcess(proc), parent=proc)

    def from_process_id(self, pid):
        return Application(self.impl.FromProcessID(pid))

    def from_hwnd(self, hwnd):
        return Application(self.impl.FromMainWindowHandle(hwnd))

    @property
    def active(self):
        return Application(self.impl.PrimaryInstance)

    @property
    def top_most(self):
        return Application(self.impl.TopMostInstance)


# noinspection PyRedeclaration
class Application(ImplCapsule, Application):
    impl: Application

    __slots__ = ImplCapsule.__slots__

    @property
    def app(self) -> Application:
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


class AppCapsule(ImplCapsule):
    @property
    def app(self) -> "Application":
        if self.parent is not None:
            return self.parent.app
        else:
            return Application(self.impl.application)


# noinspection PyRedeclaration
class Workbooks(AppCapsule, SequenceCapsule):
    impl: Workbooks
    parent: Application

    __slots__ = ImplCapsule.__slots__

    def __getitem__(self, item) -> "Workbook":
        return super().__getitem__(item)

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
                    workbook = self.impl.add()
                    workbook.save_as(str(path.absolute()))
                    return Workbook(workbook, self)
        elif path is not None:
            for workbook in self.impl:
                if os.path.abspath(workbook.full_name) == os.path.abspath(path):
                    return Workbook(workbook, self)
            else:
                return Workbook(self.impl.open(path), self)


# noinspection PyRedeclaration
class Workbook(AppCapsule, Workbook):
    impl: Workbook
    parent: Workbooks

    __slots__ = ImplCapsule.__slots__

    @property
    def sheets(self):
        return Sheets(self.impl.sheets, self)

    @property
    def active_sheet(self):
        return Worksheet(self.impl.active_sheet, self)


# noinspection PyRedeclaration
class Sheets(AppCapsule, SequenceCapsule, Sheets):
    impl: Sheets
    parent: Workbook

    __slots__ = ImplCapsule.__slots__

    def __getitem__(self, item) -> "Worksheet":
        return super().__getitem__(item)

    def wrap(self, obj):
        return Worksheet(obj, self)

    @property
    def active(self):
        return self.parent.active_sheet


# noinspection PyRedeclaration
class Worksheet(AppCapsule, Worksheet):
    impl: Worksheet
    parent: Sheets

    __slots__ = ImplCapsule.__slots__

    @property
    def cells(self):
        return Range(self.impl.cells)

    @property
    def active(self):
        selection = self.app.selection
        if selection.worksheet.name == self.impl.name:
            return Range(selection.impl, self)

        return None


# noinspection PyRedeclaration
class Range(AppCapsule, Range):
    impl: Range
    parent: Worksheet

    __slots__ = ImplCapsule.__slots__

    def __getitem__(self, key) -> "Range":
        row, col = key
        return self.impl.get(row + 1, col + 1)

    def __setitem__(self, key, value: "Range"):
        row, col = key
        self.impl.set(row + 1, col + 1, value)


def patch_2():
    ctx = globals()
    for key, value in revsered_ctx.items():
        ctx[key].__annotations__["impl"] = value


patch_2()

apps = Apps()
