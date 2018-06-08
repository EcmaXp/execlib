from typing import *
from .enums import *

__all__ = ['Action', 'Actions', 'AddIn', 'AddIns', 'AddIns2', 'Adjustments', 'AllowEditRange', 'AllowEditRanges', 'Application', 'Areas', 'Author', 'AutoCorrect', 'AutoFilter', 'AutoRecover', 'Border', 'Borders', 'CalculatedFields', 'CalculatedItems', 'CalculatedMember', 'CalculatedMembers', 'CalloutFormat', 'CellFormat', 'Characters', 'Chart', 'ChartArea', 'ChartColorFormat', 'ChartFillFormat', 'ChartFormat', 'ChartGroup', 'ChartTitle', 'ColorFormat', 'Comment', 'CommentThreaded', 'Comments', 'CommentsThreaded', 'Connections', 'ConnectorFormat', 'ControlFormat', 'Corners', 'CubeField', 'CubeFields', 'CustomProperties', 'CustomProperty', 'CustomView', 'CustomViews', 'DataFeedConnection', 'DataTable', 'DefaultPivotTableLayoutOptions', 'DefaultWebOptions', 'Diagram', 'DiagramNode', 'DiagramNodeChildren', 'DiagramNodes', 'Dialog', 'DialogFrame', 'DialogSheet', 'Dialogs', 'DisplayFormat', 'DownBars', 'DropLines', 'Error', 'ErrorCheckingOptions', 'Errors', 'FileExportConverter', 'FileExportConverters', 'FillFormat', 'Filter', 'Filters', 'Floor', 'Font', 'FormatColor', 'FormatConditions', 'FreeformBuilder', 'Graphic', 'GroupShapes', 'HPageBreak', 'HPageBreaks', 'HeaderFooter', 'HiLoLines', 'Hyperlink', 'Hyperlinks', 'Icon', 'IconSet', 'IconSets', 'Interior', 'Legend', 'LineFormat', 'LinkFormat', 'ListColumn', 'ListColumns', 'ListDataFormat', 'ListObject', 'ListObjects', 'ListRow', 'ListRows', 'Mailer', 'Menu', 'MenuBar', 'MenuBars', 'MenuItem', 'MenuItems', 'Menus', 'Model', 'Model3DFormat', 'ModelConnection', 'ModelFormatBoolean', 'ModelFormatCurrency', 'ModelFormatDate', 'ModelFormatDecimalNumber', 'ModelFormatGeneral', 'ModelFormatPercentageNumber', 'ModelFormatScientificNumber', 'ModelFormatWholeNumber', 'ModelMeasure', 'ModelMeasures', 'ModelRelationship', 'ModelRelationships', 'ModelTable', 'ModelTableColumn', 'ModelTableColumns', 'ModelTables', 'Module', 'Modules', 'MultiThreadedCalculation', 'Name', 'Names', 'ODBCConnection', 'ODBCError', 'ODBCErrors', 'OLEDBConnection', 'OLEDBError', 'OLEDBErrors', 'OLEFormat', 'Outline', 'Page', 'PageSetup', 'Pages', 'Pane', 'Panes', 'Parameter', 'Parameters', 'Phonetic', 'Phonetics', 'PictureFormat', 'PivotAxis', 'PivotCache', 'PivotCaches', 'PivotCell', 'PivotField', 'PivotFields', 'PivotFilter', 'PivotFilters', 'PivotFormula', 'PivotFormulas', 'PivotItem', 'PivotItemList', 'PivotLayout', 'PivotLine', 'PivotLineCells', 'PivotLines', 'PivotTable', 'PivotTableChangeList', 'PivotValueCell', 'PlotArea', 'ProtectedViewWindow', 'ProtectedViewWindows', 'Protection', 'PublishObject', 'PublishObjects', 'PublishedDoc', 'PublishedDocs', 'Queries', 'QueryTable', 'QueryTables', 'QuickAnalysis', 'RTD', 'Range', 'Ranges', 'RecentFile', 'RecentFiles', 'Research', 'RoutingSlip', 'SeriesLines', 'ServerViewableItems', 'ShadowFormat', 'Shape', 'ShapeNode', 'ShapeNodes', 'ShapeRange', 'Shapes', 'SheetViews', 'Sheets', 'Slicer', 'SlicerCache', 'SlicerCacheLevel', 'SlicerCacheLevels', 'SlicerCaches', 'SlicerItem', 'SlicerItems', 'SlicerPivotTables', 'Slicers', 'SmartTag', 'SmartTagAction', 'SmartTagActions', 'SmartTagOptions', 'SmartTagRecognizer', 'SmartTagRecognizers', 'SmartTags', 'Sort', 'SortField', 'SortFields', 'SoundNote', 'SparkAxes', 'SparkColor', 'SparkHorizontalAxis', 'SparkPoints', 'SparkVerticalAxis', 'Sparkline', 'SparklineGroup', 'SparklineGroups', 'Speech', 'SpellingOptions', 'Style', 'Styles', 'Tab', 'TableObject', 'TableStyle', 'TableStyleElement', 'TableStyleElements', 'TableStyles', 'TextConnection', 'TextEffectFormat', 'TextFrame', 'TextFrame2', 'ThreeDFormat', 'TickLabels', 'TimelineState', 'TimelineViewState', 'Toolbar', 'ToolbarButton', 'ToolbarButtons', 'Toolbars', 'TreeviewControl', 'UpBars', 'UsedObjects', 'UserAccess', 'UserAccessList', 'VPageBreak', 'VPageBreaks', 'Validation', 'ValueChange', 'Walls', 'Watch', 'Watches', 'WebOptions', 'Window', 'Windows', 'Workbook', 'WorkbookConnection', 'WorkbookQuery', 'Workbooks', 'Worksheet', 'WorksheetDataConnection', 'WorksheetFunction', 'XPath', 'XmlDataBinding', 'XmlMap', 'XmlMaps', 'XmlNamespace', 'XmlNamespaces', 'XmlSchema', 'XmlSchemas']

class GroupShapes:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: object) -> ShapeType: pass


    def get_range(self, index: object) -> ShapeRangeType: pass

GroupShapesType = GroupShapes


class IconSet:

    def __iter__(self, index: int) -> Iterable[IconType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    id: XlIconSet
    count: int

    def get(self, index: int) -> IconType: pass

IconSetType = IconSet


class TableStyleElements:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: XlTableStyleElementType) -> TableStyleElementType: pass

TableStyleElementsType = TableStyleElements


class SlicerItems:

    def __iter__(self, index: int) -> Iterable[SlicerItemType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, index: int) -> SlicerItemType: pass

SlicerItemsType = SlicerItems


class FormatConditions:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: object) -> object: pass


    def add(self, type: XlFormatConditionType, Operator: object = None, formula1: object = None, formula2: object = None, String: object = None, text_operator: object = None, date_operator: object = None, scope_type: object = None) -> object: pass


    def delete(self): pass


    def add_color_scale(self, color_scale_type: int) -> object: pass


    def add_databar(self) -> object: pass


    def add_icon_set_condition(self) -> object: pass


    def add_top10(self) -> object: pass


    def add_above_average(self) -> object: pass


    def add_unique_values(self) -> object: pass

FormatConditionsType = FormatConditions


class MenuBars:

    def __iter__(self, index: int) -> Iterable[MenuBarType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, name: object = None) -> MenuBarType: pass

    count: int

    def get(self, index: int) -> MenuBarType: pass

MenuBarsType = MenuBars


class OLEFormat:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def activate(self): pass

    object: object
    prog_id: str

    def verb(self, verb: object = None): pass

OLEFormatType = OLEFormat


class XmlMaps:

    def __iter__(self, index: int) -> Iterable[XmlMapType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, schema: str, root_element_name: object = None) -> XmlMapType: pass

    count: int

    def get(self, index: int) -> XmlMapType: pass

XmlMapsType = XmlMaps


class ShapeNodes:
    parent: object
    count: int

    def item(self, index: object) -> ShapeNodeType: pass


    def delete(self, index: int): pass


    def insert(self, index: int, segment_type, editing_type, x1: float, y1: float, x2: float = None, y2: float = None, x3: float = None, y3: float = None): pass


    def set_editing_type(self, index: int, editing_type): pass


    def set_position(self, index: int, x1: float, y1: float): pass


    def set_segment_type(self, index: int, segment_type): pass

ShapeNodesType = ShapeNodes


class Sheets:

    def __iter__(self, index: int) -> Iterable[WorksheetType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: WorkbookType

    def add(self, before: WorksheetType = None, after: WorksheetType = None, count: int = None, type: object = None) -> object: pass


    def copy(self, before: WorksheetType = None, after: WorksheetType = None): pass

    count: int

    def delete(self): pass


    def fill_across_sheets(self, range: RangeType, type: XlFillWith = None): pass


    def move(self, before: WorksheetType = None, after: WorksheetType = None): pass


    def print_preview(self, enable_changes: object = None): pass


    def select(self, replace: WorksheetType = None): pass

    h_page_breaks: HPageBreaksType
    v_page_breaks: VPageBreaksType
    visible: object

    def print_out(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None): pass


    def print_out_ex(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None, ignore_print_areas: object = None): pass


    def add2(self, before: object = None, after: object = None, count: object = None, new_layout: object = None) -> object: pass


    def get(self, index: int) -> WorksheetType: pass

SheetsType = Sheets


class ToolbarButton:
    application: ApplicationType
    creator: XlCreator
    parent: object
    built_in: bool
    built_in_face: bool

    def copy(self, toolbar: ToolbarType, before: int): pass


    def copy_face(self): pass


    def delete(self): pass


    def edit(self): pass

    enabled: bool
    help_context_id: int
    help_file: str
    id: int
    is_gap: bool

    def move(self, toolbar: ToolbarType, before: int): pass

    name: str
    on_action: str

    def paste_face(self): pass

    pushed: bool

    def reset(self): pass

    status_bar: str
    width: int
ToolbarButtonType = ToolbarButton


class SmartTagOptions:
    application: ApplicationType
    creator: XlCreator
    parent: object
    display_smart_tags: XlSmartTagDisplayMode
    embed_smart_tags: bool
SmartTagOptionsType = SmartTagOptions


class ValueChange:
    application: ApplicationType
    creator: XlCreator
    parent: object
    order: int
    visible_in_pivot_table: bool
    pivot_cell: PivotCellType
    tuple: str
    value: float
    allocation_value: XlAllocationValue
    allocation_method: XlAllocationMethod
    allocation_weight_expression: str

    def delete(self): pass

ValueChangeType = ValueChange


class Shape:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def apply(self): pass


    def delete(self): pass


    def duplicate(self) -> ShapeType: pass


    def flip(self, flip_cmd): pass


    def increment_left(self, increment: float): pass


    def increment_rotation(self, increment: float): pass


    def increment_top(self, increment: float): pass


    def pick_up(self): pass


    def reroute_connections(self): pass


    def scale_height(self, factor: float, relative_to_original_size, scale: object = None): pass


    def scale_width(self, factor: float, relative_to_original_size, scale: object = None): pass


    def select(self, replace: object = None): pass


    def set_shapes_default_properties(self): pass


    def ungroup(self) -> ShapeRangeType: pass


    def z_order(self, z_order_cmd): pass

    adjustments: AdjustmentsType
    text_frame: TextFrameType
    auto_shape_type: None
    callout: CalloutFormatType
    connection_site_count: int
    connector: None
    connector_format: ConnectorFormatType
    fill: FillFormatType
    group_items: GroupShapesType
    height: float
    horizontal_flip: None
    left: float
    line: LineFormatType
    lock_aspect_ratio: None
    name: str
    nodes: ShapeNodesType
    rotation: float
    picture_format: PictureFormatType
    shadow: ShadowFormatType
    text_effect: TextEffectFormatType
    three_d: ThreeDFormatType
    top: float
    type: None
    vertical_flip: None
    vertices: object
    visible: None
    width: float
    z_order_position: int
    hyperlink: HyperlinkType
    black_white_mode: None
    drawing_object: object
    on_action: str
    locked: bool
    top_left_cell: RangeType
    bottom_right_cell: RangeType
    placement: XlPlacement

    def copy(self): pass


    def cut(self): pass


    def copy_picture(self, appearance: object = None, format: object = None): pass

    control_format: ControlFormatType
    link_format: LinkFormatType
    ole_format: OLEFormatType
    form_control_type: XlFormControl
    alternative_text: str
    script: None
    diagram_node: DiagramNodeType
    has_diagram_node: None
    diagram: DiagramType
    has_diagram: None
    child: None
    parent_group: ShapeType
    canvas_items: None
    id: int

    def canvas_crop_left(self, increment: float): pass


    def canvas_crop_top(self, increment: float): pass


    def canvas_crop_right(self, increment: float): pass


    def canvas_crop_bottom(self, increment: float): pass

    chart: ChartType
    has_chart: None
    text_frame2: TextFrame2Type
    shape_style: None
    background_style: None
    soft_edge: None
    glow: None
    reflection: None
    has_smart_art: None
    smart_art: None
    title: str
    graphic_style: None
    model3d: Model3DFormatType
ShapeType = Shape


class CommentsThreaded:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: int) -> CommentThreadedType: pass

CommentsThreadedType = CommentsThreaded


class Slicer:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    caption: str
    top: float
    left: float
    disable_move_resize_ui: bool
    width: float
    height: float
    row_height: float
    column_width: float
    number_of_columns: int
    display_header: bool
    locked: bool
    slicer_cache: SlicerCacheType
    slicer_cache_level: SlicerCacheLevelType
    shape: ShapeType
    style: object

    def delete(self): pass


    def cut(self): pass


    def copy(self): pass

    active_item: SlicerItemType
    timeline_view_state: TimelineViewStateType
    slicer_cache_type: XlSlicerCacheType
SlicerType = Slicer


class PivotCell:
    application: ApplicationType
    creator: XlCreator
    parent: object
    pivot_cell_type: XlPivotCellType
    pivot_table: PivotTableType
    data_field: PivotFieldType
    pivot_field: PivotFieldType
    pivot_item: PivotItemType
    row_items: PivotItemListType
    column_items: PivotItemListType
    range: RangeType
    custom_subtotal_function: XlConsolidationFunction
    pivot_row_line: PivotLineType
    pivot_column_line: PivotLineType

    def allocate_change(self): pass


    def discard_change(self): pass

    data_source_value: object
    cell_changed: XlCellChangedState
    mdx: str
    server_actions: ActionsType
PivotCellType = PivotCell


class LineFormat:
    parent: object
    back_color: ColorFormatType
    begin_arrowhead_length: None
    begin_arrowhead_style: None
    begin_arrowhead_width: None
    dash_style: None
    end_arrowhead_length: None
    end_arrowhead_style: None
    end_arrowhead_width: None
    fore_color: ColorFormatType
    pattern: None
    style: None
    transparency: float
    visible: None
    weight: float
    inset_pen: None
LineFormatType = LineFormat


class Pages:

    def __iter__(self, index: int) -> Iterable[PageType]: pass

    count: int

    def get(self, index: int) -> PageType: pass

PagesType = Pages


class DownBars:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str

    def select(self) -> object: pass

    border: BorderType

    def delete(self) -> object: pass

    interior: InteriorType
    fill: ChartFillFormatType
    format: ChartFormatType

    def set_property(self, id: str, value: object): pass


    def get_property(self, id: str) -> object: pass

DownBarsType = DownBars


class ModelRelationships:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: object) -> ModelRelationshipType: pass


    def add(self, foreign_key_column: ModelTableColumnType, primary_key_column: ModelTableColumnType) -> ModelRelationshipType: pass


    def detect_relationships(self, pivot_table: PivotTableType): pass

ModelRelationshipsType = ModelRelationships


class ChartGroup:
    application: ApplicationType
    creator: XlCreator
    parent: object
    axis_group: XlAxisGroup
    doughnut_hole_size: int
    down_bars: DownBarsType
    drop_lines: DropLinesType
    first_slice_angle: int
    gap_width: int
    has_drop_lines: bool
    has_hi_lo_lines: bool
    has_radar_axis_labels: bool
    has_series_lines: bool
    has_up_down_bars: bool
    hi_lo_lines: HiLoLinesType
    index: int
    overlap: int
    radar_axis_labels: TickLabelsType

    def series_collection(self, index: object = None) -> object: pass

    series_lines: SeriesLinesType
    sub_type: int
    type: int
    up_bars: UpBarsType
    vary_by_categories: bool
    size_represents: XlSizeRepresents
    bubble_scale: int
    show_negative_bubbles: bool
    split_type: XlChartSplitType
    split_value: object
    second_plot_size: int
    has3d_shading: bool

    def full_category_collection(self, index: object = None) -> object: pass


    def category_collection(self, index: object = None) -> object: pass

    bins_type: XlBinsType
    bin_width_value: float
    bins_count_value: int
    bins_overflow_enabled: bool
    bins_overflow_value: float
    bins_underflow_enabled: bool
    bins_underflow_value: float
ChartGroupType = ChartGroup


class PivotValueCell:
    application: ApplicationType
    creator: XlCreator
    parent: object
    pivot_cell: PivotCellType
    value: object

    def show_detail(self): pass

    server_actions: ActionsType
PivotValueCellType = PivotValueCell


class TableStyles:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def add(self, table_style_name: str) -> TableStyleType: pass


    def item(self, index: object) -> TableStyleType: pass

TableStylesType = TableStyles


class ModelMeasure:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    associated_table: ModelTableType
    formula: str
    format_information: object
    description: str

    def delete(self): pass

ModelMeasureType = ModelMeasure


class SoundNote:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def delete(self) -> object: pass


    def import_(self, filename: str) -> object: pass


    def play(self) -> object: pass


    def record(self) -> object: pass

SoundNoteType = SoundNote


class UsedObjects:

    def __iter__(self, index: int) -> Iterable[object]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, index: int) -> object: pass

UsedObjectsType = UsedObjects


class Shapes:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: object) -> ShapeType: pass


    def add_callout(self, type, left: float, top: float, width: float, height: float) -> ShapeType: pass


    def add_connector(self, type, begin_x: float, begin_y: float, end_x: float, end_y: float) -> ShapeType: pass


    def add_curve(self, safe_array_of_points: object) -> ShapeType: pass


    def add_label(self, orientation, left: float, top: float, width: float, height: float) -> ShapeType: pass


    def add_line(self, begin_x: float, begin_y: float, end_x: float, end_y: float) -> ShapeType: pass


    def add_picture(self, filename: str, link_to_file, save_with_document, left: float, top: float, width: float, height: float) -> ShapeType: pass


    def add_polyline(self, safe_array_of_points: object) -> ShapeType: pass


    def add_shape(self, type, left: float, top: float, width: float, height: float) -> ShapeType: pass


    def add_text_effect(self, preset_text_effect, text: str, font_name: str, font_size: float, font_bold, font_italic, left: float, top: float) -> ShapeType: pass


    def add_textbox(self, orientation, left: float, top: float, width: float, height: float) -> ShapeType: pass


    def build_freeform(self, editing_type, x1: float, y1: float) -> FreeformBuilderType: pass


    def get_range(self, index: object) -> ShapeRangeType: pass


    def select_all(self): pass


    def add_form_control(self, type: XlFormControl, left: int, top: int, width: int, height: int) -> ShapeType: pass


    def add_ole_object(self, class_type: object = None, filename: object = None, link: object = None, display_as_icon: object = None, icon_file_name: object = None, icon_index: object = None, icon_label: object = None, left: object = None, top: object = None, width: object = None, height: object = None) -> ShapeType: pass


    def add_diagram(self, type, left: float, top: float, width: float, height: float) -> ShapeType: pass


    def add_canvas(self, left: float, top: float, width: float, height: float) -> ShapeType: pass


    def add_chart(self, xl_chart_type: object = None, left: object = None, top: object = None, width: object = None, height: object = None) -> ShapeType: pass


    def add_smart_art(self, layout, left: object = None, top: object = None, width: object = None, height: object = None) -> ShapeType: pass


    def add_chart2(self, style: object = None, xl_chart_type: object = None, left: object = None, top: object = None, width: object = None, height: object = None, new_layout: object = None) -> ShapeType: pass


    def add_picture2(self, filename: str, link_to_file, save_with_document, left: float, top: float, width: float, height: float, compress) -> ShapeType: pass


    def add3d_model(self, filename: str, link_to_file: object = None, save_with_document: object = None, left: object = None, top: object = None, width: object = None, height: object = None) -> ShapeType: pass

ShapesType = Shapes


class SortFields:

    def __iter__(self, index: int) -> Iterable[SortFieldType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, key: RangeType, sort_on: object = None, order: object = None, custom_order: object = None, data_option: object = None) -> SortFieldType: pass

    count: int

    def clear(self): pass


    def get(self, index: int) -> SortFieldType: pass

SortFieldsType = SortFields


class SheetViews:

    def __iter__(self, index: int) -> Iterable[object]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, index: int) -> object: pass

SheetViewsType = SheetViews


class PivotAxis:
    application: ApplicationType
    creator: XlCreator
    parent: object
    pivot_lines: PivotLinesType
PivotAxisType = PivotAxis


class PivotCache:
    application: ApplicationType
    creator: XlCreator
    parent: object
    background_query: bool
    connection: object
    enable_refresh: bool
    index: int
    memory_used: int
    optimize_cache: bool
    record_count: int

    def refresh(self): pass

    refresh_date: None
    refresh_name: str
    refresh_on_file_open: bool
    sql: object
    save_password: bool
    source_data: object
    command_text: object
    command_type: XlCmdType
    query_type: XlQueryType
    maintain_connection: bool
    refresh_period: int
    recordset: object

    def reset_timer(self): pass

    local_connection: object

    def create_pivot_table(self, table_destination: object, table_name: object = None, read_data: object = None, default_version: object = None) -> PivotTableType: pass

    use_local_connection: bool
    ado_connection: object
    is_connected: bool

    def make_connection(self): pass

    olap: bool
    source_type: XlPivotTableSourceType
    missing_items_limit: XlPivotTableMissingItems
    source_connection_file: str
    source_data_file: str
    robust_connect: XlRobustConnect

    def save_as_odc(self, odc_file_name: str, description: object = None, keywords: object = None): pass

    workbook_connection: WorkbookConnectionType
    version: XlPivotTableVersionList
    upgrade_on_refresh: bool

    def create_pivot_chart(self, chart_destination: object, xl_chart_type: object = None, left: object = None, top: object = None, width: object = None, height: object = None) -> ShapeType: pass

PivotCacheType = PivotCache


class ModelTables:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: object) -> ModelTableType: pass

ModelTablesType = ModelTables


class SmartTagRecognizer:
    application: ApplicationType
    creator: XlCreator
    parent: object
    enabled: bool
    prog_id: str
    full_name: str
SmartTagRecognizerType = SmartTagRecognizer


class ChartArea:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str

    def select(self) -> object: pass

    border: BorderType

    def clear(self) -> object: pass


    def clear_contents(self) -> object: pass


    def copy(self) -> object: pass

    font: FontType
    shadow: bool

    def clear_formats(self) -> object: pass

    height: float
    interior: InteriorType
    fill: ChartFillFormatType
    left: float
    top: float
    width: float
    auto_scale_font: object
    format: ChartFormatType
    rounded_corners: bool
ChartAreaType = ChartArea


class CalculatedItems:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: object) -> PivotItemType: pass


    def add(self, name: str, formula: str, use_standard_formula: object = None) -> PivotItemType: pass

CalculatedItemsType = CalculatedItems


class Styles:

    def __iter__(self, index: int) -> Iterable[StyleType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, name: str, based_on: object = None) -> StyleType: pass

    count: int

    def merge(self, workbook: object) -> object: pass


    def get(self, index: int) -> StyleType: pass

StylesType = Styles


class WorkbookConnection:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    description: str
    type: XlConnectionType
    oledb_connection: OLEDBConnectionType
    odbc_connection: ODBCConnectionType
    ranges: RangesType

    def delete(self): pass


    def refresh(self): pass

    model_connection: ModelConnectionType
    worksheet_data_connection: WorksheetDataConnectionType
    refresh_with_refresh_all: bool
    text_connection: TextConnectionType
    data_feed_connection: DataFeedConnectionType
    in_model: bool
    model_tables: ModelTablesType
WorkbookConnectionType = WorkbookConnection


class ModelFormatDate:
    application: ApplicationType
    creator: XlCreator
    parent: object
    format_string: str
ModelFormatDateType = ModelFormatDate


class Sparkline:
    application: ApplicationType
    creator: XlCreator
    parent: object
    location: RangeType
    source_data: str

    def modify_location(self, range: RangeType): pass


    def modify_source_data(self, formula: str): pass

SparklineType = Sparkline


class SortField:
    application: ApplicationType
    creator: XlCreator
    parent: object
    sort_on: XlSortOn
    sort_on_value: object
    key: RangeType
    order: XlSortOrder
    custom_order: object
    data_option: XlSortDataOption
    priority: int

    def delete(self): pass


    def modify_key(self, key: RangeType): pass


    def set_icon(self, icon: IconType): pass

SortFieldType = SortField


class TableStyle:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    name_local: str
    built_in: bool
    table_style_elements: TableStyleElementsType
    show_as_available_table_style: bool
    show_as_available_pivot_table_style: bool

    def delete(self): pass


    def duplicate(self, new_table_style_name: object = None) -> TableStyleType: pass

    show_as_available_slicer_style: bool
    show_as_available_timeline_style: bool
TableStyleType = TableStyle


class SlicerCaches:

    def __iter__(self, index: int) -> Iterable[SlicerCacheType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def add(self, source: object, source_field: object, name: object = None) -> SlicerCacheType: pass


    def add2(self, source: object, source_field: object, name: object = None, slicer_cache_type: object = None) -> SlicerCacheType: pass


    def get(self, index: int) -> SlicerCacheType: pass

SlicerCachesType = SlicerCaches


class AddIn:
    application: ApplicationType
    creator: XlCreator
    parent: object
    author: str
    comments: str
    full_name: str
    installed: bool
    keywords: str
    name: str
    path: str
    subject: str
    title: str
    prog_id: str
    clsid: str
    is_open: bool
AddInType = AddIn


class Queries:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def add(self, name: str, formula: str, description: object = None) -> WorkbookQueryType: pass


    def item(self, name_or_index: object) -> WorkbookQueryType: pass

    fast_combine: bool
QueriesType = Queries


class DisplayFormat:
    application: ApplicationType
    creator: XlCreator
    parent: object
    borders: BordersType

    def get_characters(self, start: object = None, length: object = None) -> CharactersType: pass

    font: FontType
    style: object
    add_indent: object
    formula_hidden: object
    horizontal_alignment: object
    indent_level: object
    interior: InteriorType
    locked: object
    merge_cells: object
    number_format: object
    number_format_local: object
    orientation: object
    reading_order: int
    shrink_to_fit: object
    vertical_alignment: object
    wrap_text: object
DisplayFormatType = DisplayFormat


class PivotField:
    application: ApplicationType
    creator: XlCreator
    parent: object
    calculation: XlPivotFieldCalculation
    child_field: PivotFieldType

    def get_child_items(self, index: object = None) -> object: pass

    current_page: object
    data_range: RangeType
    data_type: XlPivotFieldDataType
    function: XlConsolidationFunction
    group_level: object

    def get_hidden_items(self, index: object = None) -> object: pass

    label_range: RangeType
    name: str
    number_format: str
    orientation: XlPivotFieldOrientation
    show_all_items: bool
    parent_field: PivotFieldType

    def get_parent_items(self, index: object = None) -> object: pass


    def pivot_items(self, index: object = None) -> object: pass

    position: object
    source_name: str

    def get_subtotals(self, index: object = None) -> object: pass


    def set_subtotals(self, index: object = None, _param2: object = None): pass

    base_field: object
    base_item: object
    total_levels: object
    value: str

    def get_visible_items(self, index: object = None) -> object: pass


    def calculated_items(self) -> CalculatedItemsType: pass


    def delete(self): pass

    drag_to_column: bool
    drag_to_hide: bool
    drag_to_page: bool
    drag_to_row: bool
    drag_to_data: bool
    formula: str
    is_calculated: bool
    memory_used: int
    server_based: bool

    def auto_sort(self, order: int, field: str): pass


    def auto_show(self, type: int, range: int, count: int, field: str): pass

    auto_sort_order: int
    auto_sort_field: str
    auto_show_type: int
    auto_show_range: int
    auto_show_count: int
    auto_show_field: str
    layout_blank_line: bool
    layout_subtotal_location: XlSubtototalLocationType
    layout_page_break: bool
    layout_form: XlLayoutFormType
    subtotal_name: str
    caption: str
    drilled_down: bool
    cube_field: CubeFieldType
    current_page_name: str
    standard_formula: str
    hidden_items_list: object
    database_sort: bool
    is_member_property: bool
    property_parent_field: PivotFieldType
    property_order: int
    enable_item_selection: bool
    current_page_list: object

    def add_page_item(self, item: str, clear_list: object = None): pass

    hidden: bool

    def drill_to(self, field: str): pass

    use_member_property_as_caption: bool
    member_property_caption: str
    display_as_tooltip: bool
    display_in_report: bool
    display_as_caption: bool
    layout_compact_row: bool
    include_new_items_in_filter: bool
    visible_items_list: object
    pivot_filters: PivotFiltersType
    auto_sort_pivot_line: PivotLineType
    auto_sort_custom_subtotal: int
    showing_in_axis: bool
    enable_multiple_page_items: bool
    all_items_visible: bool

    def clear_manual_filter(self): pass


    def clear_all_filters(self): pass


    def clear_value_filters(self): pass


    def clear_label_filters(self): pass


    def auto_sort_ex(self, order: int, field: str, pivot_line: object = None, custom_subtotal: object = None): pass

    source_caption: str
    show_detail: bool
    repeat_labels: bool

    def auto_group(self): pass

PivotFieldType = PivotField


class ChartFormat:
    application: ApplicationType
    creator: XlCreator
    parent: object
    fill: FillFormatType
    glow: None
    line: LineFormatType
    picture_format: PictureFormatType
    shadow: ShadowFormatType
    soft_edge: None
    text_frame2: TextFrame2Type
    three_d: ThreeDFormatType
    adjustments: AdjustmentsType
    auto_shape_type: None
ChartFormatType = ChartFormat


class Research:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def query(self, service_id: str, query_string: object = None, query_language: object = None, use_selection: object = None, launch_query: object = None) -> object: pass


    def is_research_service(self, service_id: str) -> bool: pass


    def set_language_pair(self, language_from: int, language_to: int) -> object: pass

ResearchType = Research


class HeaderFooter:
    text: str
    picture: GraphicType
HeaderFooterType = HeaderFooter


class Characters:
    application: ApplicationType
    creator: XlCreator
    parent: object
    caption: str
    count: int

    def delete(self) -> object: pass

    font: FontType

    def insert(self, String: str) -> object: pass

    text: str
    phonetic_characters: str
CharactersType = Characters


class Sort:
    application: ApplicationType
    creator: XlCreator
    parent: object
    rng: RangeType
    header: XlYesNoGuess
    match_case: bool
    orientation: XlSortOrientation
    sort_method: XlSortMethod
    sort_fields: SortFieldsType

    def set_range(self, rng: RangeType): pass


    def apply(self): pass

SortType = Sort


class SparkAxes:
    application: ApplicationType
    creator: XlCreator
    parent: object
    vertical: SparkVerticalAxisType
    horizontal: SparkHorizontalAxisType
SparkAxesType = SparkAxes


class SparklineGroup:

    def __iter__(self, index: int) -> Iterable[SparklineType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int
    location: RangeType
    source_data: str
    date_range: str

    def modify_location(self, location: RangeType): pass


    def modify_source_data(self, source_data: str): pass


    def modify(self, location: RangeType, source_data: str): pass


    def modify_date_range(self, date_range: str): pass


    def delete(self): pass

    type: XlSparkType
    series_color: FormatColorType
    points: SparkPointsType
    axes: SparkAxesType
    display_blanks_as: XlDisplayBlanksAs
    display_hidden: bool
    line_weight: object
    plot_by: XlSparklineRowCol

    def get(self, index: int) -> SparklineType: pass

SparklineGroupType = SparklineGroup


class PictureFormat:
    parent: object

    def increment_brightness(self, increment: float): pass


    def increment_contrast(self, increment: float): pass

    brightness: float
    color_type: None
    contrast: float
    crop_bottom: float
    crop_left: float
    crop_right: float
    crop_top: float
    transparency_color: int
    transparent_background: None
    crop: None
PictureFormatType = PictureFormat


class XmlNamespaces:

    def __iter__(self, index: int) -> Iterable[XmlNamespaceType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int
    value: str

    def install_manifest(self, path: str, install_for_all_users: object = None): pass


    def get(self, index: int) -> XmlNamespaceType: pass

XmlNamespacesType = XmlNamespaces


class Dialog:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def show(self, arg1: object = None, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> bool: pass

DialogType = Dialog


class ODBCErrors:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: int) -> ODBCErrorType: pass

ODBCErrorsType = ODBCErrors


class Menu:
    application: ApplicationType
    creator: XlCreator
    parent: object
    caption: str

    def delete(self): pass

    enabled: bool
    index: int
    menu_items: MenuItemsType
MenuType = Menu


class Comment:
    application: ApplicationType
    creator: XlCreator
    parent: object
    author: str
    shape: ShapeType
    visible: bool

    def text(self, text: object = None, start: object = None, overwrite: object = None) -> str: pass


    def delete(self): pass


    def next(self) -> CommentType: pass


    def previous(self) -> CommentType: pass

CommentType = Comment


class HiLoLines:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str

    def select(self) -> object: pass

    border: BorderType

    def delete(self) -> object: pass

    format: ChartFormatType
HiLoLinesType = HiLoLines


class MenuItems:

    def __iter__(self, index: int) -> Iterable[object]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, caption: str, on_action: object = None, shortcut_key: object = None, before: object = None, restore: object = None, status_bar: object = None, help_file: object = None, help_context_id: object = None) -> MenuItemType: pass


    def add_menu(self, caption: str, before: object = None, restore: object = None) -> MenuType: pass

    count: int

    def get(self, index: int) -> object: pass

MenuItemsType = MenuItems


class DiagramNodes:

    def item(self, index: object) -> DiagramNodeType: pass


    def select_all(self): pass

    parent: object
    count: int
DiagramNodesType = DiagramNodes


class PivotTableChangeList:

    def __iter__(self, index: int) -> Iterable[ValueChangeType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def add(self, tuple: str, value: float, allocation_value: object = None, allocation_method: object = None, allocation_weight_expression: object = None) -> ValueChangeType: pass


    def get(self, index: int) -> ValueChangeType: pass

PivotTableChangeListType = PivotTableChangeList


class SparkColor:
    application: ApplicationType
    creator: XlCreator
    parent: object
    visible: bool
    color: FormatColorType
SparkColorType = SparkColor


class CalloutFormat:
    parent: object

    def automatic_length(self): pass


    def custom_drop(self, drop: float): pass


    def custom_length(self, length: float): pass


    def preset_drop(self, drop_type): pass

    accent: None
    angle: None
    auto_attach: None
    auto_length: None
    border: None
    drop: float
    drop_type: None
    gap: float
    length: float
    type: None
CalloutFormatType = CalloutFormat


class TimelineViewState:
    application: ApplicationType
    creator: XlCreator
    parent: object
    show_header: bool
    show_selection_label: bool
    show_time_level: bool
    show_horizontal_scrollbar: bool
    level: XlTimelineLevel
TimelineViewStateType = TimelineViewState


class TableStyleElement:
    application: ApplicationType
    creator: XlCreator
    parent: object
    has_format: bool
    interior: InteriorType
    borders: BordersType
    font: FontType
    stripe_size: int

    def clear(self): pass

TableStyleElementType = TableStyleElement


class WorksheetFunction:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def count(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def is_na(self, arg1: object) -> bool: pass


    def is_error(self, arg1: object) -> bool: pass


    def sum(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def average(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def min(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def max(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def npv(self, arg1: float, arg2: object, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def st_dev(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def dollar(self, arg1: float, arg2: object = None) -> str: pass


    def fixed(self, arg1: float, arg2: object = None, arg3: object = None) -> str: pass


    def pi(self) -> float: pass


    def ln(self, arg1: float) -> float: pass


    def log10(self, arg1: float) -> float: pass


    def round(self, arg1: float, arg2: float) -> float: pass


    def lookup(self, arg1: object, arg2: object, arg3: object = None) -> object: pass


    def index(self, arg1: object, arg2: float, arg3: object = None, arg4: object = None) -> object: pass


    def rept(self, arg1: str, arg2: float) -> str: pass


    def and_(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> bool: pass


    def or_(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> bool: pass


    def d_count(self, arg1: RangeType, arg2: object, arg3: object) -> float: pass


    def d_sum(self, arg1: RangeType, arg2: object, arg3: object) -> float: pass


    def d_average(self, arg1: RangeType, arg2: object, arg3: object) -> float: pass


    def d_min(self, arg1: RangeType, arg2: object, arg3: object) -> float: pass


    def d_max(self, arg1: RangeType, arg2: object, arg3: object) -> float: pass


    def d_st_dev(self, arg1: RangeType, arg2: object, arg3: object) -> float: pass


    def var(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def d_var(self, arg1: RangeType, arg2: object, arg3: object) -> float: pass


    def text(self, arg1: object, arg2: str) -> str: pass


    def lin_est(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None) -> object: pass


    def trend(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None) -> object: pass


    def log_est(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None) -> object: pass


    def growth(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None) -> object: pass


    def pv(self, arg1: float, arg2: float, arg3: float, arg4: object = None, arg5: object = None) -> float: pass


    def fv(self, arg1: float, arg2: float, arg3: float, arg4: object = None, arg5: object = None) -> float: pass


    def n_per(self, arg1: float, arg2: float, arg3: float, arg4: object = None, arg5: object = None) -> float: pass


    def pmt(self, arg1: float, arg2: float, arg3: float, arg4: object = None, arg5: object = None) -> float: pass


    def rate(self, arg1: float, arg2: float, arg3: float, arg4: object = None, arg5: object = None, arg6: object = None) -> float: pass


    def m_irr(self, arg1: object, arg2: float, arg3: float) -> float: pass


    def irr(self, arg1: object, arg2: object = None) -> float: pass


    def match(self, arg1: object, arg2: object, arg3: object = None) -> float: pass


    def weekday(self, arg1: object, arg2: object = None) -> float: pass


    def search(self, arg1: str, arg2: str, arg3: object = None) -> float: pass


    def transpose(self, arg1: object) -> object: pass


    def atan2(self, arg1: float, arg2: float) -> float: pass


    def asin(self, arg1: float) -> float: pass


    def acos(self, arg1: float) -> float: pass


    def choose(self, arg1: object, arg2: object, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> object: pass


    def h_lookup(self, arg1: object, arg2: object, arg3: object, arg4: object = None) -> object: pass


    def v_lookup(self, arg1: object, arg2: object, arg3: object, arg4: object = None) -> object: pass


    def log(self, arg1: float, arg2: object = None) -> float: pass


    def proper(self, arg1: str) -> str: pass


    def trim(self, arg1: str) -> str: pass


    def replace(self, arg1: str, arg2: float, arg3: float, arg4: str) -> str: pass


    def substitute(self, arg1: str, arg2: str, arg3: str, arg4: object = None) -> str: pass


    def find(self, arg1: str, arg2: str, arg3: object = None) -> float: pass


    def is_err(self, arg1: object) -> bool: pass


    def is_text(self, arg1: object) -> bool: pass


    def is_number(self, arg1: object) -> bool: pass


    def sln(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def syd(self, arg1: float, arg2: float, arg3: float, arg4: float) -> float: pass


    def ddb(self, arg1: float, arg2: float, arg3: float, arg4: float, arg5: object = None) -> float: pass


    def clean(self, arg1: str) -> str: pass


    def m_determ(self, arg1: object) -> float: pass


    def m_inverse(self, arg1: object) -> object: pass


    def m_mult(self, arg1: object, arg2: object) -> object: pass


    def ipmt(self, arg1: float, arg2: float, arg3: float, arg4: float, arg5: object = None, arg6: object = None) -> float: pass


    def ppmt(self, arg1: float, arg2: float, arg3: float, arg4: float, arg5: object = None, arg6: object = None) -> float: pass


    def count_a(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def product(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def fact(self, arg1: float) -> float: pass


    def d_product(self, arg1: RangeType, arg2: object, arg3: object) -> float: pass


    def is_non_text(self, arg1: object) -> bool: pass


    def st_dev_p(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def var_p(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def d_st_dev_p(self, arg1: RangeType, arg2: object, arg3: object) -> float: pass


    def d_var_p(self, arg1: RangeType, arg2: object, arg3: object) -> float: pass


    def is_logical(self, arg1: object) -> bool: pass


    def d_count_a(self, arg1: RangeType, arg2: object, arg3: object) -> float: pass


    def us_dollar(self, arg1: float, arg2: float) -> str: pass


    def find_b(self, arg1: str, arg2: str, arg3: object = None) -> float: pass


    def search_b(self, arg1: str, arg2: str, arg3: object = None) -> float: pass


    def replace_b(self, arg1: str, arg2: float, arg3: float, arg4: str) -> str: pass


    def round_up(self, arg1: float, arg2: float) -> float: pass


    def round_down(self, arg1: float, arg2: float) -> float: pass


    def rank(self, arg1: float, arg2: RangeType, arg3: object = None) -> float: pass


    def days360(self, arg1: object, arg2: object, arg3: object = None) -> float: pass


    def vdb(self, arg1: float, arg2: float, arg3: float, arg4: float, arg5: float, arg6: object = None, arg7: object = None) -> float: pass


    def median(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def sum_product(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def sinh(self, arg1: float) -> float: pass


    def cosh(self, arg1: float) -> float: pass


    def tanh(self, arg1: float) -> float: pass


    def asinh(self, arg1: float) -> float: pass


    def acosh(self, arg1: float) -> float: pass


    def atanh(self, arg1: float) -> float: pass


    def d_get(self, arg1: RangeType, arg2: object, arg3: object) -> object: pass


    def db(self, arg1: float, arg2: float, arg3: float, arg4: float, arg5: object = None) -> float: pass


    def frequency(self, arg1: object, arg2: object) -> object: pass


    def ave_dev(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def beta_dist(self, arg1: float, arg2: float, arg3: float, arg4: bool, arg5: object = None, arg6: object = None) -> float: pass


    def gamma_ln(self, arg1: float) -> float: pass


    def beta_inv(self, arg1: float, arg2: float, arg3: float, arg4: object = None, arg5: object = None) -> float: pass


    def binom_dist(self, arg1: float, arg2: float, arg3: float, arg4: bool) -> float: pass


    def chi_dist(self, arg1: float, arg2: float) -> float: pass


    def chi_inv(self, arg1: float, arg2: float) -> float: pass


    def combin(self, arg1: float, arg2: float) -> float: pass


    def confidence(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def crit_binom(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def even(self, arg1: float) -> float: pass


    def expon_dist(self, arg1: float, arg2: float, arg3: bool) -> float: pass


    def f_dist(self, arg1: float, arg2: float, arg3: float, arg4: bool) -> float: pass


    def f_inv(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def fisher(self, arg1: float) -> float: pass


    def fisher_inv(self, arg1: float) -> float: pass


    def floor(self, arg1: float, arg2: float) -> float: pass


    def gamma_dist(self, arg1: float, arg2: float, arg3: float, arg4: bool) -> float: pass


    def gamma_inv(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def ceiling(self, arg1: float, arg2: float) -> float: pass


    def hyp_geom_dist(self, arg1: float, arg2: float, arg3: float, arg4: float, arg5: bool) -> float: pass


    def log_norm_dist(self, arg1: float, arg2: float, arg3: float, arg4: bool) -> float: pass


    def log_inv(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def neg_binom_dist(self, arg1: float, arg2: float, arg3: float, arg4: bool) -> float: pass


    def norm_dist(self, arg1: float, arg2: float, arg3: float, arg4: bool) -> float: pass


    def norm_s_dist(self, arg1: float, arg2: bool) -> float: pass


    def norm_inv(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def norm_s_inv(self, arg1: float) -> float: pass


    def standardize(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def odd(self, arg1: float) -> float: pass


    def permut(self, arg1: float, arg2: float) -> float: pass


    def poisson(self, arg1: float, arg2: float, arg3: bool) -> float: pass


    def t_dist(self, arg1: float, arg2: float, arg3: bool) -> float: pass


    def weibull(self, arg1: float, arg2: float, arg3: float, arg4: bool) -> float: pass


    def sum_xmy2(self, arg1: object, arg2: object) -> float: pass


    def sum_x2my2(self, arg1: object, arg2: object) -> float: pass


    def sum_x2py2(self, arg1: object, arg2: object) -> float: pass


    def chi_test(self, arg1: object, arg2: object) -> float: pass


    def correl(self, arg1: object, arg2: object) -> float: pass


    def covar(self, arg1: object, arg2: object) -> float: pass


    def forecast(self, arg1: float, arg2: object, arg3: object) -> float: pass


    def f_test(self, arg1: object, arg2: object) -> float: pass


    def intercept(self, arg1: object, arg2: object) -> float: pass


    def pearson(self, arg1: object, arg2: object) -> float: pass


    def r_sq(self, arg1: object, arg2: object) -> float: pass


    def st_eyx(self, arg1: object, arg2: object) -> float: pass


    def slope(self, arg1: object, arg2: object) -> float: pass


    def t_test(self, arg1: object, arg2: object, arg3: float, arg4: float) -> float: pass


    def prob(self, arg1: object, arg2: object, arg3: float, arg4: object = None) -> float: pass


    def dev_sq(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def geo_mean(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def har_mean(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def sum_sq(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def kurt(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def skew(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def z_test(self, arg1: object, arg2: float, arg3: object = None) -> float: pass


    def large(self, arg1: object, arg2: float) -> float: pass


    def small(self, arg1: object, arg2: float) -> float: pass


    def quartile(self, arg1: object, arg2: float) -> float: pass


    def percentile(self, arg1: object, arg2: float) -> float: pass


    def percent_rank(self, arg1: object, arg2: float, arg3: object = None) -> float: pass


    def mode(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def trim_mean(self, arg1: object, arg2: float) -> float: pass


    def t_inv(self, arg1: float, arg2: float) -> float: pass


    def power(self, arg1: float, arg2: float) -> float: pass


    def radians(self, arg1: float) -> float: pass


    def degrees(self, arg1: float) -> float: pass


    def subtotal(self, arg1: float, arg2: RangeType, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def sum_if(self, arg1: RangeType, arg2: object, arg3: object = None) -> float: pass


    def count_if(self, arg1: RangeType, arg2: object) -> float: pass


    def count_blank(self, arg1: RangeType) -> float: pass


    def ispmt(self, arg1: float, arg2: float, arg3: float, arg4: float) -> float: pass


    def roman(self, arg1: float, arg2: object = None) -> str: pass


    def asc(self, arg1: str) -> str: pass


    def dbcs(self, arg1: str) -> str: pass


    def phonetic(self, arg1: RangeType) -> str: pass


    def baht_text(self, arg1: float) -> str: pass


    def thai_day_of_week(self, arg1: float) -> str: pass


    def thai_digit(self, arg1: str) -> str: pass


    def thai_month_of_year(self, arg1: float) -> str: pass


    def thai_num_sound(self, arg1: float) -> str: pass


    def thai_num_string(self, arg1: float) -> str: pass


    def thai_string_length(self, arg1: str) -> float: pass


    def is_thai_digit(self, arg1: str) -> bool: pass


    def round_baht_down(self, arg1: float) -> float: pass


    def round_baht_up(self, arg1: float) -> float: pass


    def thai_year(self, arg1: float) -> float: pass


    def rtd(self, prog_id: object, server: object, topic1: object, topic2: object = None, topic3: object = None, topic4: object = None, topic5: object = None, topic6: object = None, topic7: object = None, topic8: object = None, topic9: object = None, topic10: object = None, topic11: object = None, topic12: object = None, topic13: object = None, topic14: object = None, topic15: object = None, topic16: object = None, topic17: object = None, topic18: object = None, topic19: object = None, topic20: object = None, topic21: object = None, topic22: object = None, topic23: object = None, topic24: object = None, topic25: object = None, topic26: object = None, topic27: object = None, topic28: object = None) -> object: pass


    def hex2bin(self, arg1: object, arg2: object = None) -> str: pass


    def hex2dec(self, arg1: object) -> str: pass


    def hex2oct(self, arg1: object, arg2: object = None) -> str: pass


    def dec2bin(self, arg1: object, arg2: object = None) -> str: pass


    def dec2hex(self, arg1: object, arg2: object = None) -> str: pass


    def dec2oct(self, arg1: object, arg2: object = None) -> str: pass


    def oct2bin(self, arg1: object, arg2: object = None) -> str: pass


    def oct2hex(self, arg1: object, arg2: object = None) -> str: pass


    def oct2dec(self, arg1: object) -> str: pass


    def bin2dec(self, arg1: object) -> str: pass


    def bin2oct(self, arg1: object, arg2: object = None) -> str: pass


    def bin2hex(self, arg1: object, arg2: object = None) -> str: pass


    def im_sub(self, arg1: object, arg2: object) -> str: pass


    def im_div(self, arg1: object, arg2: object) -> str: pass


    def im_power(self, arg1: object, arg2: object) -> str: pass


    def im_abs(self, arg1: object) -> str: pass


    def im_sqrt(self, arg1: object) -> str: pass


    def im_ln(self, arg1: object) -> str: pass


    def im_log2(self, arg1: object) -> str: pass


    def im_log10(self, arg1: object) -> str: pass


    def im_sin(self, arg1: object) -> str: pass


    def im_cos(self, arg1: object) -> str: pass


    def im_exp(self, arg1: object) -> str: pass


    def im_argument(self, arg1: object) -> str: pass


    def im_conjugate(self, arg1: object) -> str: pass


    def imaginary(self, arg1: object) -> float: pass


    def im_real(self, arg1: object) -> float: pass


    def complex(self, arg1: object, arg2: object, arg3: object = None) -> str: pass


    def im_sum(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> str: pass


    def im_product(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> str: pass


    def series_sum(self, arg1: object, arg2: object, arg3: object, arg4: object) -> float: pass


    def fact_double(self, arg1: object) -> float: pass


    def sqrt_pi(self, arg1: object) -> float: pass


    def quotient(self, arg1: object, arg2: object) -> float: pass


    def delta(self, arg1: object, arg2: object = None) -> float: pass


    def ge_step(self, arg1: object, arg2: object = None) -> float: pass


    def is_even(self, arg1: object) -> bool: pass


    def is_odd(self, arg1: object) -> bool: pass


    def m_round(self, arg1: object, arg2: object) -> float: pass


    def erf(self, arg1: object, arg2: object = None) -> float: pass


    def erf_c(self, arg1: object) -> float: pass


    def bessel_j(self, arg1: object, arg2: object) -> float: pass


    def bessel_k(self, arg1: object, arg2: object) -> float: pass


    def bessel_y(self, arg1: object, arg2: object) -> float: pass


    def bessel_i(self, arg1: object, arg2: object) -> float: pass


    def xirr(self, arg1: object, arg2: object, arg3: object = None) -> float: pass


    def xnpv(self, arg1: object, arg2: object) -> float: pass


    def price_mat(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object, arg6: object = None) -> float: pass


    def yield_mat(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object, arg6: object = None) -> float: pass


    def int_rate(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object = None) -> float: pass


    def received(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object = None) -> float: pass


    def disc(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object = None) -> float: pass


    def price_disc(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object = None) -> float: pass


    def yield_disc(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object = None) -> float: pass


    def t_bill_eq(self, arg1: object, arg2: object, arg3: object = None) -> float: pass


    def t_bill_price(self, arg1: object, arg2: object, arg3: object = None) -> float: pass


    def t_bill_yield(self, arg1: object, arg2: object, arg3: object = None) -> float: pass


    def price(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object, arg6: object, arg7: object = None) -> float: pass


    def dollar_de(self, arg1: object, arg2: object) -> float: pass


    def dollar_fr(self, arg1: object, arg2: object) -> float: pass


    def nominal(self, arg1: object, arg2: object) -> float: pass


    def effect(self, arg1: object, arg2: object) -> float: pass


    def cum_princ(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object, arg6: object) -> float: pass


    def cum_i_pmt(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object, arg6: object) -> float: pass


    def e_date(self, arg1: object, arg2: object) -> float: pass


    def eo_month(self, arg1: object, arg2: object) -> float: pass


    def year_frac(self, arg1: object, arg2: object, arg3: object = None) -> float: pass


    def coup_day_bs(self, arg1: object, arg2: object, arg3: object, arg4: object = None) -> float: pass


    def coup_days(self, arg1: object, arg2: object, arg3: object, arg4: object = None) -> float: pass


    def coup_days_nc(self, arg1: object, arg2: object, arg3: object, arg4: object = None) -> float: pass


    def coup_ncd(self, arg1: object, arg2: object, arg3: object, arg4: object = None) -> float: pass


    def coup_num(self, arg1: object, arg2: object, arg3: object, arg4: object = None) -> float: pass


    def coup_pcd(self, arg1: object, arg2: object, arg3: object, arg4: object = None) -> float: pass


    def duration(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object, arg6: object = None) -> float: pass


    def m_duration(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object, arg6: object = None) -> float: pass


    def odd_l_price(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object, arg6: object, arg7: object, arg8: object = None) -> float: pass


    def odd_l_yield(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object, arg6: object, arg7: object, arg8: object = None) -> float: pass


    def odd_f_price(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object, arg6: object, arg7: object, arg8: object, arg9: object = None) -> float: pass


    def odd_f_yield(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object, arg6: object, arg7: object, arg8: object, arg9: object = None) -> float: pass


    def rand_between(self, arg1: object, arg2: object) -> float: pass


    def week_num(self, arg1: object, arg2: object = None) -> float: pass


    def amor_degrc(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object, arg6: object, arg7: object = None) -> float: pass


    def amor_linc(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object, arg6: object, arg7: object = None) -> float: pass


    def convert(self, arg1: object, arg2: object, arg3: object) -> float: pass


    def accr_int(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object, arg6: object, arg7: object = None) -> float: pass


    def accr_int_m(self, arg1: object, arg2: object, arg3: object, arg4: object, arg5: object = None) -> float: pass


    def work_day(self, arg1: object, arg2: object, arg3: object = None) -> float: pass


    def network_days(self, arg1: object, arg2: object, arg3: object = None) -> float: pass


    def gcd(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def multi_nomial(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def lcm(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def fv_schedule(self, arg1: object, arg2: object) -> float: pass


    def sum_ifs(self, arg1: RangeType, arg2: RangeType, arg3: object, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None) -> float: pass


    def count_ifs(self, arg1: RangeType, arg2: object, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def average_if(self, arg1: RangeType, arg2: object, arg3: object = None) -> float: pass


    def average_ifs(self, arg1: RangeType, arg2: RangeType, arg3: object, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None) -> float: pass


    def if_error(self, arg1: object, arg2: object) -> object: pass


    def aggregate(self, arg1: float, arg2: float, arg3: RangeType, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def confidence_norm(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def confidence_t(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def chi_sq_test(self, arg1: object, arg2: object) -> float: pass


    def covariance_p(self, arg1: object, arg2: object) -> float: pass


    def covariance_s(self, arg1: object, arg2: object) -> float: pass


    def mode_mult(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> object: pass


    def mode_sngl(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def percentile_exc(self, arg1: object, arg2: float) -> float: pass


    def percentile_inc(self, arg1: object, arg2: float) -> float: pass


    def percent_rank_exc(self, arg1: object, arg2: float, arg3: object = None) -> float: pass


    def percent_rank_inc(self, arg1: object, arg2: float, arg3: object = None) -> float: pass


    def poisson_dist(self, arg1: float, arg2: float, arg3: bool) -> float: pass


    def quartile_exc(self, arg1: object, arg2: float) -> float: pass


    def quartile_inc(self, arg1: object, arg2: float) -> float: pass


    def rank_avg(self, arg1: float, arg2: RangeType, arg3: object = None) -> float: pass


    def rank_eq(self, arg1: float, arg2: RangeType, arg3: object = None) -> float: pass


    def st_dev_s(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def t_dist_2t(self, arg1: float, arg2: float) -> float: pass


    def t_dist_rt(self, arg1: float, arg2: float) -> float: pass


    def t_inv_2t(self, arg1: float, arg2: float) -> float: pass


    def var_s(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def weibull_dist(self, arg1: float, arg2: float, arg3: float, arg4: bool) -> float: pass


    def network_days_intl(self, arg1: object, arg2: object, arg3: object = None, arg4: object = None) -> float: pass


    def work_day_intl(self, arg1: object, arg2: object, arg3: object = None, arg4: object = None) -> float: pass


    def iso_ceiling(self, arg1: float, arg2: object = None) -> float: pass


    def chi_sq_dist(self, arg1: float, arg2: float, arg3: bool) -> float: pass


    def chi_sq_dist_rt(self, arg1: float, arg2: float) -> float: pass


    def chi_sq_inv(self, arg1: float, arg2: float) -> float: pass


    def chi_sq_inv_rt(self, arg1: float, arg2: float) -> float: pass


    def f_dist_rt(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def f_inv_rt(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def log_norm_inv(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def binom_inv(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def erf_precise(self, arg1: object) -> float: pass


    def erf_c_precise(self, arg1: object) -> float: pass


    def gamma_ln_precise(self, arg1: float) -> float: pass


    def ceiling_precise(self, arg1: float, arg2: object = None) -> float: pass


    def floor_precise(self, arg1: float, arg2: object = None) -> float: pass


    def acot(self, arg1: float) -> float: pass


    def acoth(self, arg1: float) -> float: pass


    def cot(self, arg1: float) -> float: pass


    def coth(self, arg1: float) -> float: pass


    def csc(self, arg1: float) -> float: pass


    def csch(self, arg1: float) -> float: pass


    def sec(self, arg1: float) -> float: pass


    def sech(self, arg1: float) -> float: pass


    def im_cot(self, arg1: object) -> str: pass


    def im_tan(self, arg1: object) -> str: pass


    def im_csc(self, arg1: object) -> str: pass


    def im_csch(self, arg1: object) -> str: pass


    def im_sec(self, arg1: object) -> str: pass


    def im_sech(self, arg1: object) -> str: pass


    def bitand(self, arg1: float, arg2: float) -> float: pass


    def bitor(self, arg1: float, arg2: float) -> float: pass


    def bitxor(self, arg1: float, arg2: float) -> float: pass


    def bitlshift(self, arg1: float, arg2: float) -> float: pass


    def bitrshift(self, arg1: float, arg2: float) -> float: pass


    def xor(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> bool: pass


    def combina(self, arg1: float, arg2: float) -> float: pass


    def permutationa(self, arg1: float, arg2: float) -> float: pass


    def p_duration(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def base(self, arg1: float, arg2: float, arg3: object = None) -> str: pass


    def decimal(self, arg1: str, arg2: float) -> float: pass


    def days(self, arg1: object, arg2: object) -> float: pass


    def binom_dist_range(self, arg1: float, arg2: float, arg3: float, arg4: object = None) -> float: pass


    def gamma(self, arg1: float) -> float: pass


    def gauss(self, arg1: float) -> float: pass


    def phi(self, arg1: float) -> float: pass


    def skew_p(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> float: pass


    def rri(self, arg1: float, arg2: float, arg3: float) -> float: pass


    def unichar(self, arg1: float) -> str: pass


    def unicode(self, arg1: str) -> float: pass


    def munit(self, arg1: float) -> object: pass


    def arabic(self, arg1: str) -> float: pass


    def iso_week_num(self, arg1: float, arg2: object = None) -> float: pass


    def number_value(self, arg1: str, arg2: str, arg3: str) -> float: pass


    def is_formula(self, arg1: RangeType) -> bool: pass


    def if_na(self, arg1: object, arg2: object) -> object: pass


    def ceiling_math(self, arg1: float, arg2: object = None, arg3: object = None) -> float: pass


    def floor_math(self, arg1: float, arg2: object = None, arg3: object = None) -> float: pass


    def im_sinh(self, arg1: object) -> str: pass


    def im_cosh(self, arg1: object) -> str: pass


    def filter_xml(self, arg1: str, arg2: str) -> object: pass


    def web_service(self, arg1: str) -> object: pass


    def encode_url(self, arg1: str) -> object: pass


    def forecast_ets(self, arg1: float, arg2: object, arg3: object, arg4: object = None, arg5: object = None, arg6: object = None) -> float: pass


    def forecast_ets_conf_int(self, arg1: float, arg2: object, arg3: object, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None) -> float: pass


    def forecast_ets_seasonality(self, arg1: object, arg2: object, arg3: object = None, arg4: object = None) -> float: pass


    def forecast_linear(self, arg1: float, arg2: object, arg3: object) -> float: pass


    def forecast_ets_stat(self, arg1: object, arg2: object, arg3: float, arg4: object = None, arg5: object = None, arg6: object = None) -> float: pass


    def max_ifs(self, arg1: RangeType, arg2: RangeType, arg3: object, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None) -> float: pass


    def min_ifs(self, arg1: RangeType, arg2: RangeType, arg3: object, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None) -> float: pass


    def text_join(self, arg1: str, arg2: bool, arg3: str, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None) -> str: pass


    def concat(self, arg1: str, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None) -> str: pass


    def sort(self, arg1: object, arg2: object = None, arg3: object = None, arg4: object = None) -> object: pass


    def unique(self, arg1: object, arg2: object = None, arg3: object = None) -> object: pass

WorksheetFunctionType = WorksheetFunction


class MultiThreadedCalculation:
    application: ApplicationType
    creator: XlCreator
    parent: object
    enabled: bool
    thread_mode: XlThreadMode
    thread_count: int
MultiThreadedCalculationType = MultiThreadedCalculation


class Menus:

    def __iter__(self, index: int) -> Iterable[MenuType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, caption: str, before: object = None, restore: object = None) -> MenuType: pass

    count: int

    def get(self, index: int) -> MenuType: pass

MenusType = Menus


class Model3DFormat:
    parent: object
    auto_fit: None
    rotation_x: float
    rotation_y: float
    rotation_z: float
    field_of_view: float
    camera_position_x: float
    camera_position_y: float
    camera_position_z: float
    look_at_point_x: float
    look_at_point_y: float
    look_at_point_z: float

    def reset_model(self, reset_size: bool = None): pass


    def increment_rotation_x(self, increment: float): pass


    def increment_rotation_y(self, increment: float): pass


    def increment_rotation_z(self, increment: float): pass

Model3DFormatType = Model3DFormat


class AutoFilter:
    application: ApplicationType
    creator: XlCreator
    parent: object
    range: RangeType
    filters: FiltersType
    filter_mode: bool
    sort: SortType

    def apply_filter(self): pass


    def show_all_data(self): pass

AutoFilterType = AutoFilter


class ListObjects:

    def __iter__(self, index: int) -> Iterable[ListObjectType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, source_type: XlListObjectSourceType = None, source: object = None, link_source: object = None, xl_list_object_has_headers: XlYesNoGuess = None, destination: object = None) -> ListObjectType: pass

    count: int

    def add_ex(self, source_type: XlListObjectSourceType = None, source: object = None, link_source: object = None, xl_list_object_has_headers: XlYesNoGuess = None, destination: object = None, table_style_name: object = None) -> ListObjectType: pass


    def get(self, index: int) -> ListObjectType: pass

ListObjectsType = ListObjects


class Author:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    provider_id: str
    user_id: str
AuthorType = Author


class DialogSheet:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def activate(self): pass


    def copy(self, before: object = None, after: object = None): pass


    def delete(self): pass

    code_name: str
    index: int

    def move(self, before: object = None, after: object = None): pass

    name: str
    next: object
    on_double_click: str
    on_sheet_activate: str
    on_sheet_deactivate: str
    page_setup: PageSetupType
    previous: object

    def print_preview(self, enable_changes: object = None): pass

    protect_contents: bool
    protect_drawing_objects: bool
    protection_mode: bool
    protect_scenarios: bool

    def select(self, replace: object = None): pass


    def unprotect(self, password: object = None): pass

    visible: XlSheetVisibility
    shapes: ShapesType

    def arcs(self, index: object = None) -> object: pass


    def buttons(self, index: object = None) -> object: pass

    enable_calculation: bool

    def chart_objects(self, index: object = None) -> object: pass


    def check_boxes(self, index: object = None) -> object: pass


    def check_spelling(self, custom_dictionary: object = None, ignore_uppercase: object = None, always_suggest: object = None, spell_lang: object = None): pass

    display_automatic_page_breaks: bool

    def drawings(self, index: object = None) -> object: pass


    def drawing_objects(self, index: object = None) -> object: pass


    def drop_downs(self, index: object = None) -> object: pass

    enable_auto_filter: bool
    enable_selection: XlEnableSelection
    enable_outlining: bool
    enable_pivot_table: bool

    def evaluate(self, name: object) -> object: pass


    def reset_all_page_breaks(self): pass


    def group_boxes(self, index: object = None) -> object: pass


    def group_objects(self, index: object = None) -> object: pass


    def labels(self, index: object = None) -> object: pass


    def lines(self, index: object = None) -> object: pass


    def list_boxes(self, index: object = None) -> object: pass

    names: NamesType

    def ole_objects(self, index: object = None) -> object: pass


    def option_buttons(self, index: object = None) -> object: pass


    def ovals(self, index: object = None) -> object: pass


    def paste(self, destination: object = None, link: object = None): pass


    def pictures(self, index: object = None) -> object: pass


    def rectangles(self, index: object = None) -> object: pass

    scroll_area: str

    def scroll_bars(self, index: object = None) -> object: pass


    def spinners(self, index: object = None) -> object: pass


    def text_boxes(self, index: object = None) -> object: pass

    h_page_breaks: HPageBreaksType
    v_page_breaks: VPageBreaksType
    query_tables: QueryTablesType
    display_page_breaks: bool
    comments: CommentsType
    hyperlinks: HyperlinksType

    def clear_circles(self): pass


    def circle_invalid(self): pass

    auto_filter: AutoFilterType
    display_right_to_left: bool
    scripts: None

    def print_out(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None): pass

    tab: TabType
    mail_envelope: None

    def save_as(self, filename: str, file_format: object = None, password: object = None, write_res_password: object = None, read_only_recommended: object = None, create_backup: object = None, add_to_mru: object = None, text_codepage: object = None, text_visual_layout: object = None, local: object = None): pass

    custom_properties: CustomPropertiesType
    smart_tags: SmartTagsType
    protection: ProtectionType

    def paste_special(self, format: object = None, link: object = None, display_as_icon: object = None, icon_file_name: object = None, icon_index: object = None, icon_label: object = None, no_html_formatting: object = None): pass


    def protect(self, password: object = None, drawing_objects: object = None, contents: object = None, scenarios: object = None, user_interface_only: object = None, allow_formatting_cells: object = None, allow_formatting_columns: object = None, allow_formatting_rows: object = None, allow_inserting_columns: object = None, allow_inserting_rows: object = None, allow_inserting_hyperlinks: object = None, allow_deleting_columns: object = None, allow_deleting_rows: object = None, allow_sorting: object = None, allow_filtering: object = None, allow_using_pivot_tables: object = None): pass


    def print_out_ex(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None): pass

    enable_format_conditions_calculation: bool
    sort: SortType

    def export_as_fixed_format(self, type: XlFixedFormatType, filename: object = None, quality: object = None, include_doc_properties: object = None, ignore_print_areas: object = None, from_: object = None, to: object = None, open_after_publish: object = None, fixed_format_ext_class_ptr: object = None): pass

    printed_comment_pages: int

    def export_as_fixed_format2(self, type: XlFixedFormatType, filename: object = None, quality: object = None, include_doc_properties: object = None, ignore_print_areas: object = None, from_: object = None, to: object = None, open_after_publish: object = None, fixed_format_ext_class_ptr: object = None, work_identity: object = None): pass


    def save_as2(self, filename: str, file_format: object = None, password: object = None, write_res_password: object = None, read_only_recommended: object = None, create_backup: object = None, add_to_mru: object = None, text_codepage: object = None, text_visual_layout: object = None, local: object = None): pass

    comments_threaded: CommentsThreadedType
    default_button: object
    dialog_frame: DialogFrameType

    def edit_boxes(self, index: object = None) -> object: pass

    focus: object

    def hide(self, cancel: object = None) -> bool: pass


    def show(self) -> bool: pass

DialogSheetType = DialogSheet


class TextFrame:
    application: ApplicationType
    creator: XlCreator
    parent: object
    margin_bottom: float
    margin_left: float
    margin_right: float
    margin_top: float
    orientation: None

    def characters(self, start: object = None, length: object = None) -> CharactersType: pass

    horizontal_alignment: XlHAlign
    vertical_alignment: XlVAlign
    auto_size: bool
    reading_order: int
    auto_margins: bool
    vertical_overflow: XlOartVerticalOverflow
    horizontal_overflow: XlOartHorizontalOverflow
TextFrameType = TextFrame


class ListObject:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def delete(self): pass


    def publish(self, target: object, link_source: bool) -> str: pass


    def refresh(self): pass


    def unlink(self): pass


    def unlist(self): pass


    def update_changes(self, i_conflict_type: XlListConflict = None): pass


    def resize(self, range: RangeType): pass

    active: bool
    data_body_range: RangeType
    display_right_to_left: bool
    header_row_range: RangeType
    insert_row_range: RangeType
    list_columns: ListColumnsType
    list_rows: ListRowsType
    name: str
    query_table: QueryTableType
    range: RangeType
    show_auto_filter: bool
    show_totals: bool
    source_type: XlListObjectSourceType
    totals_row_range: RangeType
    share_point_url: str
    xml_map: XmlMapType
    display_name: str
    show_headers: bool
    auto_filter: AutoFilterType
    table_style: object
    show_table_style_first_column: bool
    show_table_style_last_column: bool
    show_table_style_row_stripes: bool
    show_table_style_column_stripes: bool
    sort: SortType
    comment: str

    def export_to_visio(self): pass

    alternative_text: str
    summary: str
    table_object: TableObjectType
    slicers: SlicersType
    show_auto_filter_drop_down: bool
ListObjectType = ListObject


class ListRow:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def delete(self): pass

    index: int
    invalid_data: bool
    range: RangeType
ListRowType = ListRow


class OLEDBErrors:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: int) -> OLEDBErrorType: pass

OLEDBErrorsType = OLEDBErrors


class Watches:

    def __iter__(self, index: int) -> Iterable[WatchType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, source: object) -> WatchType: pass

    count: int

    def delete(self): pass


    def get(self, index: int) -> WatchType: pass

WatchesType = Watches


class IconSets:

    def __iter__(self, index: int) -> Iterable[object]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, index: int) -> object: pass

IconSetsType = IconSets


class Speech:

    def speak(self, text: str, speak_async: object = None, speak_xml: object = None, purge: object = None): pass

    direction: XlSpeakDirection
    speak_cell_on_enter: bool
SpeechType = Speech


class PivotLineCells:

    def __iter__(self, index: int) -> Iterable[PivotCellType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int
    full: bool

    def get(self, index: int) -> PivotCellType: pass

PivotLineCellsType = PivotLineCells


class ODBCError:
    application: ApplicationType
    creator: XlCreator
    parent: object
    sql_state: str
    error_string: str
ODBCErrorType = ODBCError


class PageSetup:
    application: ApplicationType
    creator: XlCreator
    parent: object
    black_and_white: bool
    bottom_margin: float
    center_footer: str
    center_header: str
    center_horizontally: bool
    center_vertically: bool
    chart_size: XlObjectSize
    draft: bool
    first_page_number: int
    fit_to_pages_tall: object
    fit_to_pages_wide: object
    footer_margin: float
    header_margin: float
    left_footer: str
    left_header: str
    left_margin: float
    order: XlOrder
    orientation: XlPageOrientation
    paper_size: XlPaperSize
    print_area: str
    print_gridlines: bool
    print_headings: bool
    print_notes: bool

    def get_print_quality(self, index: object = None) -> object: pass


    def set_print_quality(self, index: object = None, _param2: object = None): pass

    print_title_columns: str
    print_title_rows: str
    right_footer: str
    right_header: str
    right_margin: float
    top_margin: float
    zoom: object
    print_comments: XlPrintLocation
    print_errors: XlPrintErrors
    center_header_picture: GraphicType
    center_footer_picture: GraphicType
    left_header_picture: GraphicType
    left_footer_picture: GraphicType
    right_header_picture: GraphicType
    right_footer_picture: GraphicType
    odd_and_even_pages_header_footer: bool
    different_first_page_header_footer: bool
    scale_with_doc_header_footer: bool
    align_margins_header_footer: bool
    pages: PagesType
    even_page: PageType
    first_page: PageType
PageSetupType = PageSetup


class Floor:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str

    def select(self) -> object: pass

    border: BorderType

    def clear_formats(self) -> object: pass

    interior: InteriorType
    fill: ChartFillFormatType
    picture_type: object

    def paste(self): pass

    thickness: int
    format: ChartFormatType
FloorType = Floor


class PivotLine:
    application: ApplicationType
    creator: XlCreator
    parent: object
    line_type: XlPivotLineType
    position: int
    pivot_line_cells: PivotLineCellsType
    pivot_line_cells_full: PivotLineCellsType
PivotLineType = PivotLine


class Icon:
    application: ApplicationType
    creator: XlCreator
    parent: IconSetType
    index: int
IconType = Icon


class TimelineState:
    application: ApplicationType
    creator: XlCreator
    parent: object
    start_date: object
    end_date: object
    filter_type: XlPivotFilterType
    filter_value1: object
    filter_value2: object
    single_range_filter_state: bool

    def set_filter_date_range(self, start_date: object, end_date: object) -> XlFilterStatus: pass

TimelineStateType = TimelineState


class Watch:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def delete(self): pass

    source: object
WatchType = Watch


class ModelFormatBoolean:
    application: ApplicationType
    creator: XlCreator
    parent: object
ModelFormatBooleanType = ModelFormatBoolean


class Workbook:
    application: ApplicationType
    creator: XlCreator
    parent: WorkbookType
    accept_labels_in_formulas: bool

    def activate(self): pass

    active_chart: ChartType
    active_sheet: WorksheetType
    author: str
    auto_update_frequency: int
    auto_update_save_changes: bool
    change_history_duration: int
    builtin_document_properties: object

    def change_file_access(self, mode: XlFileAccess, write_password: object = None, notify: object = None): pass


    def change_link(self, name: str, new_name: str, type: XlLinkType = None): pass

    charts: SheetsType

    def close(self, save_changes: object = None, filename: object = None, route_workbook: object = None): pass

    code_name: str

    def get_colors(self, index: object = None) -> object: pass


    def set_colors(self, index: object = None, rhs: object = None): pass

    command_bars: None
    comments: str
    conflict_resolution: XlSaveConflictResolution
    container: object
    create_backup: bool
    custom_document_properties: object
    date1904: bool

    def delete_number_format(self, number_format: str): pass

    dialog_sheets: SheetsType
    display_drawing_objects: XlDisplayDrawingObjects

    def exclusive_access(self) -> bool: pass

    file_format: XlFileFormat

    def forward_mailer(self): pass

    full_name: str
    has_mailer: bool
    has_password: bool
    has_routing_slip: bool
    is_addin: bool
    keywords: str

    def link_info(self, name: str, link_info: XlLinkInfo, type: object = None, edition_ref: object = None) -> object: pass


    def link_sources(self, type: object = None) -> object: pass

    mailer: MailerType

    def merge_workbook(self, filename: object): pass

    modules: SheetsType
    multi_user_editing: bool
    name: str
    names: NamesType

    def new_window(self) -> WindowType: pass

    on_save: str
    on_sheet_activate: str
    on_sheet_deactivate: str

    def open_links(self, name: str, read_only: object = None, type: object = None): pass

    path: str
    personal_view_list_settings: bool
    personal_view_print_settings: bool

    def pivot_caches(self) -> PivotCachesType: pass


    def post(self, dest_name: object = None): pass

    precision_as_displayed: bool

    def print_preview(self, enable_changes: object = None): pass


    def protect_sharing(self, filename: object = None, password: object = None, write_res_password: object = None, read_only_recommended: object = None, create_backup: object = None, sharing_password: object = None): pass

    protect_structure: bool
    protect_windows: bool
    read_only: bool

    def refresh_all(self): pass


    def reply(self): pass


    def reply_all(self): pass


    def remove_user(self, index: int): pass

    revision_number: int

    def route(self): pass

    routed: bool
    routing_slip: RoutingSlipType

    def run_auto_macros(self, which: XlRunAutoMacro): pass


    def save(self): pass


    def save_copy_as(self, filename: object = None): pass

    saved: bool
    save_link_values: bool

    def send_mail(self, recipients: object, subject: object = None, return_receipt: object = None): pass


    def send_mailer(self, file_format: object = None, priority: XlPriority = None): pass


    def set_link_on_data(self, name: str, procedure: object = None): pass

    sheets: SheetsType
    show_conflict_history: bool
    styles: StylesType
    subject: str
    title: str

    def unprotect(self, password: object = None): pass


    def unprotect_sharing(self, sharing_password: object = None): pass


    def update_from_file(self): pass


    def update_link(self, name: object = None, type: object = None): pass

    update_remote_references: bool
    user_control: bool
    user_status: object
    custom_views: CustomViewsType
    windows: WindowsType
    worksheets: SheetsType
    write_reserved: bool
    write_reserved_by: str
    excel4intl_macro_sheets: SheetsType
    excel4macro_sheets: SheetsType
    template_remove_ext_data: bool

    def highlight_changes_options(self, when: object = None, who: object = None, where: object = None): pass

    highlight_changes_on_screen: bool
    keep_change_history: bool
    list_changes_on_new_sheet: bool

    def purge_change_history_now(self, days: int, sharing_password: object = None): pass


    def accept_all_changes(self, when: object = None, who: object = None, where: object = None): pass


    def reject_all_changes(self, when: object = None, who: object = None, where: object = None): pass


    def pivot_table_wizard(self, source_type: object = None, source_data: object = None, table_destination: object = None, table_name: object = None, row_grand: object = None, column_grand: object = None, save_data: object = None, has_auto_format: object = None, auto_page: object = None, reserved: object = None, background_query: object = None, optimize_cache: object = None, page_field_order: object = None, page_field_wrap_count: object = None, read_data: object = None, connection: object = None): pass


    def reset_colors(self): pass

    vb_project: None

    def follow_hyperlink(self, address: str, sub_address: object = None, new_window: object = None, add_history: object = None, extra_info: object = None, method: object = None, header_info: object = None): pass


    def add_to_favorites(self): pass

    is_inplace: bool

    def print_out(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None): pass


    def web_page_preview(self): pass

    publish_objects: PublishObjectsType
    web_options: WebOptionsType

    def reload_as(self, encoding): pass

    html_project: None
    envelope_visible: bool
    calculation_version: int

    def sblt(self, s: str): pass

    vba_signed: bool
    show_pivot_table_field_list: bool
    update_links: XlUpdateLinks

    def break_link(self, name: str, type: XlLinkType): pass


    def save_as(self, filename: object = None, file_format: object = None, password: object = None, write_res_password: object = None, read_only_recommended: object = None, create_backup: object = None, access_mode: XlSaveAsAccessMode = None, conflict_resolution: object = None, add_to_mru: object = None, text_codepage: object = None, text_visual_layout: object = None, local: object = None): pass

    enable_auto_recover: bool
    remove_personal_information: bool
    full_name_url_encoded: str

    def check_in(self, save_changes: object = None, comments: object = None, make_public: object = None): pass


    def can_check_in(self) -> bool: pass


    def send_for_review(self, recipients: object = None, subject: object = None, show_message: object = None, include_attachment: object = None): pass


    def reply_with_changes(self, show_message: object = None): pass


    def end_review(self): pass

    password: str
    write_password: str
    password_encryption_provider: str
    password_encryption_algorithm: str
    password_encryption_key_length: int

    def set_password_encryption_options(self, password_encryption_provider: object = None, password_encryption_algorithm: object = None, password_encryption_key_length: object = None, password_encryption_file_properties: object = None): pass

    password_encryption_file_properties: bool
    read_only_recommended: bool

    def protect(self, password: object = None, structure: object = None, windows: object = None): pass

    smart_tag_options: SmartTagOptionsType

    def recheck_smart_tags(self): pass

    permission: None
    shared_workspace: None
    sync: None

    def send_fax_over_internet(self, recipients: object = None, subject: object = None, show_message: object = None): pass

    xml_namespaces: XmlNamespacesType
    xml_maps: XmlMapsType

    def xml_import(self, url: str, import_map: XmlMapType, overwrite: object = None, destination: object = None) -> XlXmlImportResult: pass

    smart_document: None
    document_library_versions: None
    inactive_list_border_visible: bool
    display_ink_comments: bool

    def xml_import_xml(self, data: str, import_map: XmlMapType, overwrite: object = None, destination: object = None) -> XlXmlImportResult: pass


    def save_as_xml_data(self, filename: str, map: XmlMapType): pass


    def toggle_forms_design(self): pass

    content_type_properties: None
    connections: ConnectionsType

    def remove_document_information(self, remove_doc_info_type: XlRemoveDocInfoType): pass

    signatures: None

    def check_in_with_version(self, save_changes: object = None, comments: object = None, make_public: object = None, version_type: object = None): pass

    server_policy: None

    def lock_server_file(self): pass

    document_inspectors: None

    def get_workflow_tasks(self): pass


    def get_workflow_templates(self): pass


    def print_out_ex(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None, ignore_print_areas: object = None): pass

    server_viewable_items: ServerViewableItemsType
    table_styles: TableStylesType
    default_table_style: object
    default_pivot_table_style: object
    check_compatibility: bool
    has_vb_project: bool
    custom_xml_parts: None
    final: bool
    research: ResearchType
    theme: None

    def apply_theme(self, filename: str): pass

    excel8compatibility_mode: bool
    connections_disabled: bool

    def enable_connections(self): pass

    show_pivot_chart_active_fields: bool

    def export_as_fixed_format(self, type: XlFixedFormatType, filename: object = None, quality: object = None, include_doc_properties: object = None, ignore_print_areas: object = None, from_: object = None, to: object = None, open_after_publish: object = None, fixed_format_ext_class_ptr: object = None): pass

    icon_sets: IconSetsType
    encryption_provider: str
    do_not_prompt_for_convert: bool
    force_full_calculation: bool

    def protect_sharing_ex(self, filename: object = None, password: object = None, write_res_password: object = None, read_only_recommended: object = None, create_backup: object = None, sharing_password: object = None, file_format: object = None): pass

    slicer_caches: SlicerCachesType
    active_slicer: SlicerType
    default_slicer_style: object
    accuracy_version: int
    case_sensitive: bool
    use_whole_cell_criteria: bool
    use_wildcards: bool
    pivot_tables: object
    model: ModelType
    chart_data_point_track: bool
    default_timeline_style: object
    queries: QueriesType

    def create_forecast_sheet(self, timeline: RangeType, values: RangeType, forecast_start: object = None, forecast_end: object = None, conf_int: object = None, seasonality: object = None, data_completion: object = None, aggregation: object = None, chart_type: object = None, show_stats_table: object = None): pass

    work_identity: str

    def save_as2(self, filename: object = None, file_format: object = None, password: object = None, write_res_password: object = None, read_only_recommended: object = None, create_backup: object = None, access_mode: XlSaveAsAccessMode = None, conflict_resolution: object = None, add_to_mru: object = None, text_codepage: object = None, text_visual_layout: object = None, local: object = None, work_identity: object = None): pass


    def export_as_fixed_format2(self, type: XlFixedFormatType, filename: object = None, quality: object = None, include_doc_properties: object = None, ignore_print_areas: object = None, from_: object = None, to: object = None, open_after_publish: object = None, fixed_format_ext_class_ptr: object = None, work_identity: object = None): pass


    def publish_to_docs(self, title: str, disclosure_scope: XlPublishToDocsDisclosureScope, overwrite_url: object = None) -> str: pass


    def look_up_in_docs(self, filename: object = None) -> PublishedDocsType: pass


    def publish_to_pbi(self, publish_type: object = None, name_conflict: object = None, bstr_group_name: object = None) -> str: pass

    auto_save_on: bool

    def upgrade_comments(self): pass

WorkbookType = Workbook


class Hyperlinks:

    def __iter__(self, index: int) -> Iterable[HyperlinkType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, anchor: object, address: str, sub_address: object = None, screen_tip: object = None, text_to_display: object = None) -> object: pass

    count: int

    def delete(self): pass


    def get(self, index: int) -> HyperlinkType: pass

HyperlinksType = Hyperlinks


class PublishedDoc:
    application: ApplicationType
    creator: XlCreator
    parent: object
    title: str
    disclosure_scope: int
    url: str
    published_date: None
PublishedDocType = PublishedDoc


class CellFormat:
    application: ApplicationType
    creator: XlCreator
    parent: object
    borders: BordersType
    font: FontType
    interior: InteriorType
    number_format: object
    number_format_local: object
    add_indent: object
    indent_level: object
    horizontal_alignment: object
    vertical_alignment: object
    orientation: object
    shrink_to_fit: object
    wrap_text: object
    locked: object
    formula_hidden: object
    merge_cells: object

    def clear(self): pass

CellFormatType = CellFormat


class XmlSchema:
    application: ApplicationType
    creator: XlCreator
    parent: object
    namespace: XmlNamespaceType
    xml: str
    name: str
XmlSchemaType = XmlSchema


class Page:
    left_header: HeaderFooterType
    center_header: HeaderFooterType
    right_header: HeaderFooterType
    left_footer: HeaderFooterType
    center_footer: HeaderFooterType
    right_footer: HeaderFooterType
PageType = Page


class ColorFormat:
    parent: object
    scheme_color: int
    type: None
    tint_and_shade: float
    object_theme_color: None
    brightness: float
ColorFormatType = ColorFormat


class CubeField:
    application: ApplicationType
    creator: XlCreator
    parent: object
    cube_field_type: XlCubeFieldType
    name: str
    value: str
    orientation: XlPivotFieldOrientation
    position: int
    treeview_control: TreeviewControlType
    drag_to_column: bool
    drag_to_hide: bool
    drag_to_page: bool
    drag_to_row: bool
    drag_to_data: bool
    hidden_levels: int
    has_member_properties: bool
    layout_form: XlLayoutFormType
    pivot_fields: PivotFieldsType

    def add_member_property_field(self, property: str, property_order: object = None): pass

    enable_multiple_page_items: bool
    layout_subtotal_location: XlSubtototalLocationType
    show_in_field_list: bool

    def delete(self): pass


    def add_member_property_field_ex(self, property: str, property_order: object = None, property_displayed_in: object = None): pass

    include_new_items_in_filter: bool
    cube_field_sub_type: XlCubeFieldSubType
    all_items_visible: bool

    def clear_manual_filter(self): pass


    def create_pivot_fields(self): pass

    current_page_name: str
    is_date: bool
    caption: str
    flatten_hierarchies: bool
    hierarchize_distinct: bool

    def auto_group(self, orientation: object = None, position: object = None): pass

CubeFieldType = CubeField


class TreeviewControl:
    application: ApplicationType
    creator: XlCreator
    parent: object
    hidden: object
    drilled: object
TreeviewControlType = TreeviewControl


class SlicerItem:
    application: ApplicationType
    creator: XlCreator
    parent: SlicerCacheType
    caption: str
    name: str
    source_name: object
    source_name_standard: str
    value: str
    selected: bool
    has_data: bool
SlicerItemType = SlicerItem


class Panes:

    def __iter__(self, index: int) -> Iterable[PaneType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, index: int) -> PaneType: pass

PanesType = Panes


class Style:
    application: ApplicationType
    creator: XlCreator
    parent: object
    add_indent: bool
    built_in: bool
    borders: BordersType

    def delete(self) -> object: pass

    font: FontType
    formula_hidden: bool
    horizontal_alignment: XlHAlign
    include_alignment: bool
    include_border: bool
    include_font: bool
    include_number: bool
    include_patterns: bool
    include_protection: bool
    indent_level: int
    interior: InteriorType
    locked: bool
    merge_cells: object
    name: str
    name_local: str
    number_format: str
    number_format_local: str
    orientation: XlOrientation
    shrink_to_fit: bool
    value: str
    vertical_alignment: XlVAlign
    wrap_text: bool
    reading_order: int
StyleType = Style


class SparkPoints:
    application: ApplicationType
    creator: XlCreator
    parent: object
    negative: SparkColorType
    markers: SparkColorType
    highpoint: SparkColorType
    lowpoint: SparkColorType
    firstpoint: SparkColorType
    lastpoint: SparkColorType
SparkPointsType = SparkPoints


class MenuItem:
    application: ApplicationType
    creator: XlCreator
    parent: object
    caption: str
    checked: bool

    def delete(self): pass

    enabled: bool
    help_context_id: int
    help_file: str
    index: int
    on_action: str
    status_bar: str
MenuItemType = MenuItem


class QuickAnalysis:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def show(self, xl_quick_analysis_mode: XlQuickAnalysisMode = None): pass


    def hide(self, xl_quick_analysis_mode: XlQuickAnalysisMode = None): pass

QuickAnalysisType = QuickAnalysis


class Names:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, name: object = None, refers_to: object = None, visible: object = None, macro_type: object = None, shortcut_key: object = None, category: object = None, name_local: object = None, refers_to_local: object = None, category_local: object = None, refers_to_r1c1: object = None, refers_to_r1c1local: object = None) -> NameType: pass


    def item(self, index: object = None, index_local: object = None, refers_to: object = None) -> NameType: pass

    count: int
NamesType = Names


class Phonetics:

    def __iter__(self, index: int) -> Iterable[object]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int
    start: int
    length: int
    visible: bool
    character_type: int
    alignment: int
    font: FontType

    def delete(self): pass


    def add(self, start: int, length: int, text: str): pass

    text: str

    def get(self, index: int) -> object: pass

PhoneticsType = Phonetics


class ShadowFormat:
    parent: object

    def increment_offset_x(self, increment: float): pass


    def increment_offset_y(self, increment: float): pass

    fore_color: ColorFormatType
    obscured: None
    offset_x: float
    offset_y: float
    transparency: float
    type: None
    visible: None
    style: None
    blur: float
    size: float
    rotate_with_shape: None
ShadowFormatType = ShadowFormat


class XmlNamespace:
    application: ApplicationType
    creator: XlCreator
    parent: object
    uri: str
    prefix: str
XmlNamespaceType = XmlNamespace


class UserAccessList:

    def __iter__(self, index: int) -> Iterable[UserAccessType]: pass

    count: int

    def add(self, name: str, allow_edit: bool) -> UserAccessType: pass


    def delete_all(self): pass


    def get(self, index: int) -> UserAccessType: pass

UserAccessListType = UserAccessList


class CustomViews:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, view_name: object) -> CustomViewType: pass


    def add(self, view_name: str, print_settings: object = None, row_col_settings: object = None) -> CustomViewType: pass

CustomViewsType = CustomViews


class ModelTableColumn:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    data_type: int
ModelTableColumnType = ModelTableColumn


class ListColumn:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def delete(self): pass

    list_data_format: ListDataFormatType
    index: int
    name: str
    range: RangeType
    totals_calculation: XlTotalsCalculation
    x_path: XPathType
    share_point_formula: str
    data_body_range: RangeType
    total: RangeType
ListColumnType = ListColumn


class ControlFormat:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def add_item(self, text: str, index: object = None): pass


    def remove_all_items(self): pass


    def remove_item(self, index: int, count: object = None): pass

    drop_down_lines: int
    enabled: bool
    large_change: int
    linked_cell: str

    def get_list(self, index: object = None) -> object: pass

    list_count: int
    list_fill_range: str
    list_index: int
    locked_text: bool
    max: int
    min: int
    multi_select: int
    print_object: bool
    small_change: int
    value: int
ControlFormatType = ControlFormat


class PivotItem:
    application: ApplicationType
    creator: XlCreator
    parent: PivotFieldType

    def get_child_items(self, index: object = None) -> object: pass

    data_range: RangeType
    label_range: RangeType
    name: str
    parent_item: PivotItemType
    parent_show_detail: bool
    position: int
    show_detail: bool
    source_name: object
    value: str
    visible: bool

    def delete(self): pass

    is_calculated: bool
    record_count: int
    formula: str
    caption: str
    drilled_down: bool
    standard_formula: str
    source_name_standard: str

    def drill_to(self, field: str): pass

PivotItemType = PivotItem


class WorkbookQuery:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    formula: str
    description: str

    def delete(self): pass

WorkbookQueryType = WorkbookQuery


class SmartTagActions:

    def __iter__(self, index: int) -> Iterable[SmartTagActionType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, index: int) -> SmartTagActionType: pass

SmartTagActionsType = SmartTagActions


class ModelTable:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    source_name: str
    model_table_columns: ModelTableColumnsType
    source_workbook_connection: WorkbookConnectionType

    def refresh(self): pass

    record_count: int
ModelTableType = ModelTable


class XmlSchemas:

    def __iter__(self, index: int) -> Iterable[XmlSchemaType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, index: int) -> XmlSchemaType: pass

XmlSchemasType = XmlSchemas


class ChartFillFormat:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def one_color_gradient(self, style, variant: int, degree: float): pass


    def two_color_gradient(self, style, variant: int): pass


    def preset_textured(self, preset_texture): pass


    def solid(self): pass


    def patterned(self, pattern): pass


    def user_picture(self, picture_file: object = None, picture_format: object = None, picture_stack_unit: object = None, picture_placement: object = None): pass


    def user_textured(self, texture_file: str): pass


    def preset_gradient(self, style, variant: int, preset_gradient_type): pass

    back_color: ChartColorFormatType
    fore_color: ChartColorFormatType
    gradient_color_type: None
    gradient_degree: float
    gradient_style: None
    gradient_variant: int
    pattern: None
    preset_gradient_type: None
    preset_texture: None
    texture_name: str
    texture_type: None
    type: None
    visible: None
ChartFillFormatType = ChartFillFormat


class AddIns2:

    def __iter__(self, index: int) -> Iterable[AddInType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, filename: str, copy_file: object = None) -> AddInType: pass

    count: int

    def get(self, index: int) -> AddInType: pass

AddIns2Type = AddIns2


class ModelFormatGeneral:
    application: ApplicationType
    creator: XlCreator
    parent: object
ModelFormatGeneralType = ModelFormatGeneral


class RoutingSlip:
    application: ApplicationType
    creator: XlCreator
    parent: object
    delivery: XlRoutingSlipDelivery
    message: object

    def get_recipients(self, index: object = None) -> object: pass


    def set_recipients(self, index: object = None, _param2: object = None): pass


    def reset(self) -> object: pass

    return_when_done: bool
    status: XlRoutingSlipStatus
    subject: object
    track_status: bool
RoutingSlipType = RoutingSlip


class ModelMeasures:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: object) -> ModelMeasureType: pass


    def add(self, measure_name: str, associated_table: ModelTableType, formula: str, format_information: object, description: object = None) -> ModelMeasureType: pass

ModelMeasuresType = ModelMeasures


class FileExportConverter:
    application: ApplicationType
    creator: XlCreator
    parent: object
    extensions: str
    description: str
    file_format: int
FileExportConverterType = FileExportConverter


class SlicerPivotTables:

    def __iter__(self, index: int) -> Iterable[PivotTableType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def add_pivot_table(self, pivot_table: PivotTableType): pass


    def remove_pivot_table(self, pivot_table: object): pass


    def get(self, index: int) -> PivotTableType: pass

SlicerPivotTablesType = SlicerPivotTables


class Filter:
    application: ApplicationType
    creator: XlCreator
    parent: object
    on: bool
    criteria1: object
    criteria2: object
    operator: XlAutoFilterOperator
    count: int
FilterType = Filter


class CommentThreaded:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def add_reply(self, text: object = None) -> CommentThreadedType: pass


    def delete(self): pass


    def text(self, text: object = None, start: object = None, overwrite: object = None) -> str: pass

    supports_replies: bool
    replies: CommentsThreadedType
    author: AuthorType
    date: object

    def next(self) -> CommentThreadedType: pass


    def previous(self) -> CommentThreadedType: pass

CommentThreadedType = CommentThreaded


class SmartTag:
    application: ApplicationType
    creator: XlCreator
    parent: object
    download_url: str
    name: str
    xml: str
    range: RangeType

    def delete(self): pass

    smart_tag_actions: SmartTagActionsType
    properties: CustomPropertiesType
SmartTagType = SmartTag


class PublishObjects:

    def __iter__(self, index: int) -> Iterable[PublishObjectType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, source_type: XlSourceType, filename: str, sheet: object = None, source: object = None, html_type: object = None, div_id: object = None, title: object = None) -> PublishObjectType: pass

    count: int

    def delete(self): pass


    def publish(self): pass


    def get(self, index: int) -> PublishObjectType: pass

PublishObjectsType = PublishObjects


class Name:
    application: ApplicationType
    creator: XlCreator
    parent: object
    index: int
    category: str
    category_local: str

    def delete(self): pass

    macro_type: XlXLMMacroType
    name: str
    refers_to: object
    shortcut_key: str
    value: str
    visible: bool
    name_local: str
    refers_to_local: object
    refers_to_r1c1: object
    refers_to_r1c1local: object
    refers_to_range: RangeType
    comment: str
    workbook_parameter: bool
    valid_workbook_parameter: bool
NameType = Name


class DataTable:
    application: ApplicationType
    creator: XlCreator
    parent: object
    show_legend_key: bool
    has_border_horizontal: bool
    has_border_vertical: bool
    has_border_outline: bool
    border: BorderType
    font: FontType

    def select(self): pass


    def delete(self): pass

    auto_scale_font: object
    format: ChartFormatType
DataTableType = DataTable


class WorksheetDataConnection:
    application: ApplicationType
    creator: XlCreator
    parent: object
    connection: object
    command_text: object
    command_type: XlCmdType
WorksheetDataConnectionType = WorksheetDataConnection


class HPageBreaks:

    def __iter__(self, index: int) -> Iterable[HPageBreakType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def add(self, before: object) -> HPageBreakType: pass


    def get(self, index: int) -> HPageBreakType: pass

HPageBreaksType = HPageBreaks


class Errors:

    def __iter__(self, index: int) -> Iterable[ErrorType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def get(self, index: int) -> ErrorType: pass

ErrorsType = Errors


class Workbooks:

    def __iter__(self, index: int) -> Iterable[WorkbookType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: ApplicationType

    def add(self, template: object = None) -> WorkbookType: pass


    def close(self): pass

    count: int

    def open(self, filename: str, update_links: object = None, read_only: object = None, format: object = None, password: object = None, write_res_password: object = None, ignore_read_only_recommended: object = None, origin: object = None, delimiter: object = None, editable: object = None, notify: object = None, converter: object = None, add_to_mru: object = None, local: object = None, corrupt_load: object = None) -> WorkbookType: pass


    def open_text(self, filename: str, origin: object = None, start_row: object = None, data_type: object = None, text_qualifier: XlTextQualifier = None, consecutive_delimiter: object = None, tab: object = None, semicolon: object = None, comma: object = None, space: object = None, other: object = None, other_char: object = None, field_info: object = None, text_visual_layout: object = None, decimal_separator: object = None, thousands_separator: object = None, trailing_minus_numbers: object = None, local: object = None): pass


    def open_database(self, filename: str, command_text: object = None, command_type: object = None, background_query: object = None, import_data_as: object = None) -> WorkbookType: pass


    def check_out(self, filename: str): pass


    def can_check_out(self, filename: str) -> bool: pass


    def open_xml(self, filename: str, stylesheets: object = None, load_option: object = None) -> WorkbookType: pass


    def get(self, index: int) -> WorkbookType: pass

WorkbooksType = Workbooks


class PivotTable:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def add_fields(self, row_fields: object = None, column_fields: object = None, page_fields: object = None, add_to_table: object = None) -> object: pass


    def get_column_fields(self, index: object = None) -> object: pass

    column_grand: bool
    column_range: RangeType

    def show_pages(self, page_field: object = None) -> object: pass

    data_body_range: RangeType

    def get_data_fields(self, index: object = None) -> object: pass

    data_label_range: RangeType
    has_auto_format: bool

    def get_hidden_fields(self, index: object = None) -> object: pass

    inner_detail: str
    name: str

    def get_page_fields(self, index: object = None) -> object: pass

    page_range: RangeType
    page_range_cells: RangeType

    def pivot_fields(self, index: object = None) -> object: pass

    refresh_date: None
    refresh_name: str

    def refresh_table(self) -> bool: pass


    def get_row_fields(self, index: object = None) -> object: pass

    row_grand: bool
    row_range: RangeType
    save_data: bool
    source_data: object
    table_range1: RangeType
    table_range2: RangeType
    value: str

    def get_visible_fields(self, index: object = None) -> object: pass

    cache_index: int

    def calculated_fields(self) -> CalculatedFieldsType: pass

    display_error_string: bool
    display_null_string: bool
    enable_drilldown: bool
    enable_field_dialog: bool
    enable_wizard: bool
    error_string: str

    def get_data(self, name: str) -> float: pass


    def list_formulas(self): pass

    manual_update: bool
    merge_labels: bool
    null_string: str

    def pivot_cache(self) -> PivotCacheType: pass

    pivot_formulas: PivotFormulasType

    def pivot_table_wizard(self, source_type: object = None, source_data: object = None, table_destination: object = None, table_name: object = None, row_grand: object = None, column_grand: object = None, save_data: object = None, has_auto_format: object = None, auto_page: object = None, reserved: object = None, background_query: object = None, optimize_cache: object = None, page_field_order: object = None, page_field_wrap_count: object = None, read_data: object = None, connection: object = None): pass

    subtotal_hidden_page_items: bool
    page_field_order: int
    page_field_style: str
    page_field_wrap_count: int
    preserve_formatting: bool
    pivot_selection: str
    selection_mode: XlPTSelectionMode
    table_style: str
    tag: str

    def update(self): pass

    vacated_style: str

    def format(self, format: XlPivotFormatType): pass

    print_titles: bool
    cube_fields: CubeFieldsType
    grand_total_name: str
    small_grid: bool
    repeat_items_on_each_printed_page: bool
    totals_annotation: bool

    def pivot_select(self, name: str, mode: XlPTSelectionMode = None, use_standard_name: object = None): pass

    pivot_selection_standard: str

    def get_pivot_data(self, data_field: object = None, field1: object = None, item1: object = None, field2: object = None, item2: object = None, field3: object = None, item3: object = None, field4: object = None, item4: object = None, field5: object = None, item5: object = None, field6: object = None, item6: object = None, field7: object = None, item7: object = None, field8: object = None, item8: object = None, field9: object = None, item9: object = None, field10: object = None, item10: object = None, field11: object = None, item11: object = None, field12: object = None, item12: object = None, field13: object = None, item13: object = None, field14: object = None, item14: object = None) -> RangeType: pass

    data_pivot_field: PivotFieldType
    enable_data_value_editing: bool

    def add_data_field(self, field: object, caption: object = None, function: object = None) -> PivotFieldType: pass

    mdx: str
    view_calculated_members: bool
    calculated_members: CalculatedMembersType
    display_immediate_items: bool
    enable_field_list: bool
    visual_totals: bool
    show_page_multiple_item_label: bool
    version: XlPivotTableVersionList

    def create_cube_file(self, file: str, measures: object = None, levels: object = None, members: object = None, properties: object = None) -> str: pass

    display_empty_row: bool
    display_empty_column: bool
    show_cell_background_from_olap: bool
    pivot_column_axis: PivotAxisType
    pivot_row_axis: PivotAxisType
    show_drill_indicators: bool
    print_drill_indicators: bool
    display_member_property_tooltips: bool
    display_context_tooltips: bool

    def clear_table(self): pass

    compact_row_indent: int
    layout_row_default: XlLayoutRowType
    display_field_captions: bool

    def row_axis_layout(self, row_layout: XlLayoutRowType): pass


    def subtotal_location(self, location: XlSubtototalLocationType): pass

    active_filters: PivotFiltersType
    in_grid_drop_zones: bool

    def clear_all_filters(self): pass

    table_style2: object
    show_table_style_last_column: bool
    show_table_style_row_stripes: bool
    show_table_style_column_stripes: bool
    show_table_style_row_headers: bool
    show_table_style_column_headers: bool

    def convert_to_formulas(self, convert_filters: bool): pass

    allow_multiple_filters: bool
    compact_layout_row_header: str
    compact_layout_column_header: str
    field_list_sort_ascending: bool
    sort_using_custom_lists: bool

    def change_connection(self, conn: WorkbookConnectionType): pass


    def change_pivot_cache(self, pivot_cache: object): pass

    location: str
    enable_writeback: bool
    allocation: XlAllocation
    allocation_value: XlAllocationValue
    allocation_method: XlAllocationMethod
    allocation_weight_expression: str

    def allocate_changes(self): pass


    def commit_changes(self): pass


    def discard_changes(self): pass


    def refresh_data_source_values(self): pass


    def repeat_all_labels(self, repeat: XlPivotFieldRepeatLabels): pass

    change_list: PivotTableChangeListType
    slicers: SlicersType
    alternative_text: str
    summary: str
    visual_totals_for_sets: bool
    show_values_row: bool
    calculated_members_in_filters: bool

    def pivot_value_cell(self, rowline: object = None, columnline: object = None) -> PivotValueCellType: pass

    hidden: bool
    pivot_chart: ShapeType

    def drill_down(self, pivot_item: PivotItemType, pivot_line: object = None): pass


    def drill_up(self, pivot_item: PivotItemType, pivot_line: object = None, level_unique_name: object = None): pass


    def drill_to(self, pivot_item: PivotItemType, cube_field: CubeFieldType, pivot_line: object = None): pass


    def apply_layout(self): pass

PivotTableType = PivotTable


class SmartTagRecognizers:

    def __iter__(self, index: int) -> Iterable[SmartTagRecognizerType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int
    recognize: bool

    def get(self, index: int) -> SmartTagRecognizerType: pass

SmartTagRecognizersType = SmartTagRecognizers


class Modules:

    def __iter__(self, index: int) -> Iterable[object]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, before: object = None, after: object = None, count: object = None) -> ModuleType: pass


    def copy(self, before: object = None, after: object = None): pass

    count: int

    def delete(self): pass


    def move(self, before: object = None, after: object = None): pass


    def select(self, replace: object = None): pass

    h_page_breaks: HPageBreaksType
    v_page_breaks: VPageBreaksType
    visible: object

    def print_out(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None): pass


    def print_out_ex(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None, ignore_print_areas: object = None): pass


    def add2(self, before: object = None, after: object = None, count: object = None, new_layout: object = None) -> object: pass


    def get(self, index: int) -> object: pass

ModulesType = Modules


class ModelConnection:
    application: ApplicationType
    creator: XlCreator
    parent: object
    command_text: object
    command_type: XlCmdType
    ado_connection: object
    calculated_members: CalculatedMembersType
ModelConnectionType = ModelConnection


class Borders:

    def __iter__(self, index: XlBordersIndex) -> Iterable[BorderType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    color: object
    color_index: object
    count: int
    line_style: object
    value: object
    weight: object
    theme_color: object
    tint_and_shade: object

    def get(self, index: XlBordersIndex) -> BorderType: pass

BordersType = Borders


class Connections:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def add_from_file(self, filename: str) -> WorkbookConnectionType: pass


    def add(self, name: str, description: str, connection_string: object, command_text: object, l_cmdtype: object = None) -> WorkbookConnectionType: pass


    def item(self, index: object) -> WorkbookConnectionType: pass


    def add2(self, name: str, description: str, connection_string: object, command_text: object, l_cmdtype: object = None, create_model_connection: object = None, import_relationships: object = None) -> WorkbookConnectionType: pass


    def add_from_file2(self, filename: str, create_model_connection: object = None, import_relationships: object = None) -> WorkbookConnectionType: pass

ConnectionsType = Connections


class DefaultPivotTableLayoutOptions:
    application: ApplicationType
    creator: XlCreator
    parent: object
    in_grid_drop_zones: bool
    display_field_captions: bool
    field_list_sort_ascending: bool
    view_calculated_members: bool
    display_context_tooltips: bool
    show_drill_indicators: bool
    display_empty_column: bool
    display_empty_row: bool
    display_member_property_tooltips: bool
    show_values_row: bool
    display_null_string: bool
    null_string: str
    display_error_string: bool
    error_string: str
    has_auto_format: bool
    page_field_order: bool
    merge_labels: bool
    preserve_formatting: bool
    page_field_wrap_count: int
    compact_row_indent: int
    print_drill_indicators: bool
    repeat_items_on_each_printed_page: bool
    print_titles: bool
    allow_multiple_filters: bool
    calculated_members_in_filters: bool
    visual_totals_for_sets: bool
    visual_totals: bool
    totals_annotation: bool
    row_grand: bool
    column_grand: bool
    subtotal_hidden_page_items: bool
    sort_using_custom_lists: bool
    save_data: bool
    enable_drilldown: bool
    refresh_on_file_open: bool
    xl_missing_items_none: int
    subtotals: bool
    subtotal_location: bool
    layout_blank_line: bool
    row_axis_layout: XlLayoutRowType
    repeat_all_labels: XlPivotFieldRepeatLabels
    display_immediate_items: bool
    enable_writeback: bool
DefaultPivotTableLayoutOptionsType = DefaultPivotTableLayoutOptions


class ChartTitle:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str

    def select(self) -> object: pass

    border: BorderType

    def delete(self) -> object: pass

    interior: InteriorType
    fill: ChartFillFormatType
    caption: str

    def get_characters(self, start: object = None, length: object = None) -> CharactersType: pass

    font: FontType
    horizontal_alignment: object
    left: float
    orientation: object
    shadow: bool
    text: str
    top: float
    vertical_alignment: object
    reading_order: int
    auto_scale_font: object
    include_in_layout: bool
    position: XlChartElementPosition
    format: ChartFormatType
    height: float
    width: float
    formula: str
    formula_r1c1: str
    formula_local: str
    formula_r1c1local: str

    def set_property(self, id: str, value: object): pass


    def get_property(self, id: str) -> object: pass

ChartTitleType = ChartTitle


class Adjustments:
    parent: object
    count: int
AdjustmentsType = Adjustments


class FillFormat:
    parent: object

    def background(self): pass


    def one_color_gradient(self, style, variant: int, degree: float): pass


    def patterned(self, pattern): pass


    def preset_gradient(self, style, variant: int, preset_gradient_type): pass


    def preset_textured(self, preset_texture): pass


    def solid(self): pass


    def two_color_gradient(self, style, variant: int): pass


    def user_picture(self, picture_file: str): pass


    def user_textured(self, texture_file: str): pass

    back_color: ColorFormatType
    fore_color: ColorFormatType
    gradient_color_type: None
    gradient_degree: float
    gradient_style: None
    gradient_variant: int
    pattern: None
    preset_gradient_type: None
    preset_texture: None
    texture_name: str
    texture_type: None
    transparency: float
    type: None
    visible: None
    gradient_stops: None
    texture_offset_x: float
    texture_offset_y: float
    texture_alignment: None
    texture_horizontal_scale: float
    texture_vertical_scale: float
    texture_tile: None
    rotate_with_object: None
    picture_effects: None
    gradient_angle: float
FillFormatType = FillFormat


class CalculatedFields:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: object) -> PivotFieldType: pass


    def add(self, name: str, formula: str, use_standard_formula: object = None) -> PivotFieldType: pass

CalculatedFieldsType = CalculatedFields


class Parameter:
    application: ApplicationType
    creator: XlCreator
    parent: object
    data_type: XlParameterDataType
    type: XlParameterType
    prompt_string: str
    value: object
    source_range: RangeType
    name: str

    def set_param(self, type: XlParameterType, value: object): pass

    refresh_on_change: bool
ParameterType = Parameter


class ListDataFormat:
    application: ApplicationType
    creator: XlCreator
    parent: object
    choices: object
    decimal_places: int
    default_value: object
    is_percent: bool
    lcid: int
    max_characters: int
    max_number: object
    min_number: object
    required: bool
    type: XlListDataType
    read_only: bool
    allow_fill_in: bool
ListDataFormatType = ListDataFormat


class SmartTagAction:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str

    def execute(self): pass

    type: XlSmartTagControlType
    present_in_pane: bool
    expand_help: bool
    checkbox_state: bool
    textbox_text: str
    list_selection: int
    radio_group_selection: int
    active_x_control: object
SmartTagActionType = SmartTagAction


class MenuBar:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def activate(self): pass

    built_in: bool
    caption: str

    def delete(self): pass

    index: int
    menus: MenusType

    def reset(self): pass

MenuBarType = MenuBar


class Chart:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def activate(self): pass


    def copy(self, before: object = None, after: object = None): pass


    def delete(self): pass

    code_name: str
    index: int

    def move(self, before: object = None, after: object = None): pass

    name: str
    next: object
    on_double_click: str
    on_sheet_activate: str
    on_sheet_deactivate: str
    page_setup: PageSetupType
    previous: object

    def print_preview(self, enable_changes: object = None): pass

    protect_contents: bool
    protect_drawing_objects: bool
    protection_mode: bool

    def select(self, replace: object = None): pass


    def unprotect(self, password: object = None): pass

    visible: XlSheetVisibility
    shapes: ShapesType

    def arcs(self, index: object = None) -> object: pass

    area3d_group: ChartGroupType

    def area_groups(self, index: object = None) -> object: pass


    def auto_format(self, gallery: int, format: object = None): pass

    auto_scaling: bool

    def axes(self, type: object = None, axis_group: XlAxisGroup = None) -> object: pass


    def set_background_picture(self, filename: str): pass

    bar3d_group: ChartGroupType

    def bar_groups(self, index: object = None) -> object: pass


    def buttons(self, index: object = None) -> object: pass

    chart_area: ChartAreaType

    def chart_groups(self, index: object = None) -> object: pass


    def chart_objects(self, index: object = None) -> object: pass

    chart_title: ChartTitleType

    def chart_wizard(self, source: object = None, gallery: object = None, format: object = None, plot_by: object = None, category_labels: object = None, series_labels: object = None, has_legend: object = None, title: object = None, category_title: object = None, value_title: object = None, extra_title: object = None): pass


    def check_boxes(self, index: object = None) -> object: pass


    def check_spelling(self, custom_dictionary: object = None, ignore_uppercase: object = None, always_suggest: object = None, spell_lang: object = None): pass

    column3d_group: ChartGroupType

    def column_groups(self, index: object = None) -> object: pass


    def copy_picture(self, appearance: XlPictureAppearance = None, format: XlCopyPictureFormat = None, size: XlPictureAppearance = None): pass

    corners: CornersType

    def create_publisher(self, edition: object = None, appearance: XlPictureAppearance = None, size: XlPictureAppearance = None, contains_pict: object = None, contains_biff: object = None, contains_rtf: object = None, contains_valu: object = None): pass

    data_table: DataTableType
    depth_percent: int

    def deselect(self): pass

    display_blanks_as: XlDisplayBlanksAs

    def doughnut_groups(self, index: object = None) -> object: pass


    def drawings(self, index: object = None) -> object: pass


    def drawing_objects(self, index: object = None) -> object: pass


    def drop_downs(self, index: object = None) -> object: pass

    elevation: int

    def evaluate(self, name: object) -> object: pass

    floor: FloorType
    gap_depth: int

    def group_boxes(self, index: object = None) -> object: pass


    def group_objects(self, index: object = None) -> object: pass


    def get_has_axis(self, index1: object = None, index2: object = None) -> object: pass


    def set_has_axis(self, index1: object = None, index2: object = None, rhs: object = None): pass

    has_data_table: bool
    has_legend: bool
    has_title: bool
    height_percent: int
    hyperlinks: HyperlinksType

    def labels(self, index: object = None) -> object: pass

    legend: LegendType
    line3d_group: ChartGroupType

    def line_groups(self, index: object = None) -> object: pass


    def lines(self, index: object = None) -> object: pass


    def list_boxes(self, index: object = None) -> object: pass


    def location(self, where: XlChartLocation, name: object = None) -> ChartType: pass


    def ole_objects(self, index: object = None) -> object: pass


    def option_buttons(self, index: object = None) -> object: pass


    def ovals(self, index: object = None) -> object: pass


    def paste(self, type: object = None): pass

    perspective: int

    def pictures(self, index: object = None) -> object: pass

    pie3d_group: ChartGroupType

    def pie_groups(self, index: object = None) -> object: pass

    plot_area: PlotAreaType
    plot_visible_only: bool

    def radar_groups(self, index: object = None) -> object: pass


    def rectangles(self, index: object = None) -> object: pass

    right_angle_axes: object
    rotation: object

    def scroll_bars(self, index: object = None) -> object: pass


    def series_collection(self, index: object = None) -> object: pass

    size_with_window: bool
    show_window: bool

    def spinners(self, index: object = None) -> object: pass

    sub_type: int
    surface_group: ChartGroupType

    def text_boxes(self, index: object = None) -> object: pass

    type: int
    chart_type: XlChartType

    def apply_custom_type(self, chart_type: XlChartType, type_name: object = None): pass

    walls: WallsType
    walls_and_gridlines2d: bool

    def xy_groups(self, index: object = None) -> object: pass

    bar_shape: XlBarShape
    plot_by: XlRowCol

    def copy_chart_build(self): pass

    protect_formatting: bool
    protect_data: bool
    protect_goal_seek: bool
    protect_selection: bool

    def get_chart_element(self, x: int, y: int, element_id: int, arg1: int, arg2: int): pass


    def set_source_data(self, source: RangeType, plot_by: object = None): pass


    def export(self, filename: str, filter_name: object = None, interactive: object = None) -> bool: pass


    def refresh(self): pass

    pivot_layout: PivotLayoutType
    has_pivot_fields: bool
    scripts: None

    def print_out(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None): pass

    tab: TabType
    mail_envelope: None

    def apply_data_labels(self, type: XlDataLabelsType = None, legend_key: object = None, auto_text: object = None, has_leader_lines: object = None, show_series_name: object = None, show_category_name: object = None, show_value: object = None, show_percentage: object = None, show_bubble_size: object = None, separator: object = None): pass


    def save_as(self, filename: str, file_format: object = None, password: object = None, write_res_password: object = None, read_only_recommended: object = None, create_backup: object = None, add_to_mru: object = None, text_codepage: object = None, text_visual_layout: object = None, local: object = None): pass


    def protect(self, password: object = None, drawing_objects: object = None, contents: object = None, scenarios: object = None, user_interface_only: object = None): pass


    def apply_layout(self, layout: int, chart_type: object = None): pass


    def set_element(self, element): pass

    show_data_labels_over_maximum: bool
    side_wall: WallsType
    back_wall: WallsType

    def print_out_ex(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None): pass


    def apply_chart_template(self, filename: str): pass


    def save_chart_template(self, filename: str): pass


    def set_default_chart(self, name: object): pass


    def export_as_fixed_format(self, type: XlFixedFormatType, filename: object = None, quality: object = None, include_doc_properties: object = None, ignore_print_areas: object = None, from_: object = None, to: object = None, open_after_publish: object = None, fixed_format_ext_class_ptr: object = None): pass

    chart_style: object

    def clear_to_match_style(self): pass

    printed_comment_pages: int
    show_report_filter_field_buttons: bool
    show_legend_field_buttons: bool
    show_axis_field_buttons: bool
    show_value_field_buttons: bool
    show_all_field_buttons: bool

    def full_series_collection(self, index: object = None) -> object: pass

    category_label_level: XlCategoryLabelLevel
    series_name_level: XlSeriesNameLevel
    has_hidden_content: bool

    def delete_hidden_content(self): pass

    chart_color: object

    def clear_to_match_color_style(self): pass

    show_expand_collapse_entire_field_buttons: bool

    def export_as_fixed_format2(self, type: XlFixedFormatType, filename: object = None, quality: object = None, include_doc_properties: object = None, ignore_print_areas: object = None, from_: object = None, to: object = None, open_after_publish: object = None, fixed_format_ext_class_ptr: object = None, work_identity: object = None): pass


    def set_property(self, id: str, value: object): pass


    def get_property(self, id: str) -> object: pass

    display_value_not_available_as_blank: bool

    def save_as2(self, filename: str, file_format: object = None, password: object = None, write_res_password: object = None, read_only_recommended: object = None, create_backup: object = None, add_to_mru: object = None, text_codepage: object = None, text_visual_layout: object = None, local: object = None): pass

ChartType = Chart


class DefaultWebOptions:
    application: ApplicationType
    creator: XlCreator
    parent: object
    rely_on_css: bool
    save_hidden_data: bool
    load_pictures: bool
    organize_in_folder: bool
    update_links_on_save: bool
    use_long_file_names: bool
    check_if_office_is_html_editor: bool
    download_components: bool
    rely_on_vml: bool
    allow_png: bool
    screen_size: None
    pixels_per_inch: int
    location_of_components: str
    encoding: None
    always_save_in_default_encoding: bool
    fonts: None
    folder_suffix: str
    target_browser: None
    save_new_web_pages_as_web_archives: bool
DefaultWebOptionsType = DefaultWebOptions


class ConnectorFormat:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def begin_connect(self, connected_shape: ShapeType, connection_site: int): pass


    def begin_disconnect(self): pass


    def end_connect(self, connected_shape: ShapeType, connection_site: int): pass


    def end_disconnect(self): pass

    begin_connected: None
    begin_connected_shape: ShapeType
    begin_connection_site: int
    end_connected: None
    end_connected_shape: ShapeType
    end_connection_site: int
    type: None
ConnectorFormatType = ConnectorFormat


class CubeFields:

    def __iter__(self, index: int) -> Iterable[CubeFieldType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def add_set(self, name: str, caption: str) -> CubeFieldType: pass


    def get_measure(self, attribute_hierarchy: object, function: XlConsolidationFunction, caption: object = None) -> CubeFieldType: pass


    def get(self, index: int) -> CubeFieldType: pass

CubeFieldsType = CubeFields


class PlotArea:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str

    def select(self) -> object: pass

    border: BorderType

    def clear_formats(self) -> object: pass

    height: float
    interior: InteriorType
    fill: ChartFillFormatType
    left: float
    top: float
    width: float
    inside_left: float
    inside_top: float
    inside_width: float
    inside_height: float
    position: XlChartElementPosition
    format: ChartFormatType

    def set_property(self, id: str, value: object): pass


    def get_property(self, id: str) -> object: pass

PlotAreaType = PlotArea


class OLEDBConnection:
    application: ApplicationType
    creator: XlCreator
    parent: object
    ado_connection: object
    background_query: bool

    def cancel_refresh(self): pass

    command_text: object
    command_type: XlCmdType
    connection: object
    enable_refresh: bool
    local_connection: object
    maintain_connection: bool

    def make_connection(self): pass


    def refresh(self): pass

    refresh_date: None
    refreshing: bool
    refresh_on_file_open: bool
    refresh_period: int
    robust_connect: XlRobustConnect

    def save_as_odc(self, odc_file_name: str, description: object = None, keywords: object = None): pass

    save_password: bool
    source_connection_file: str
    source_data_file: str
    olap: bool
    use_local_connection: bool
    max_drillthrough_records: int
    is_connected: bool
    server_credentials_method: XlCredentialsMethod
    server_sso_application_id: str
    always_use_connection_file: bool
    server_fill_color: bool
    server_font_style: bool
    server_number_format: bool
    server_text_color: bool
    retrieve_in_office_ui_lang: bool

    def reconnect(self): pass

    calculated_members: CalculatedMembersType
    locale_id: int
OLEDBConnectionType = OLEDBConnection


class PivotLines:

    def __iter__(self, index: int) -> Iterable[PivotLineType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, index: int) -> PivotLineType: pass

PivotLinesType = PivotLines


class CustomProperty:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    value: object

    def delete(self): pass

CustomPropertyType = CustomProperty


class TextConnection:
    application: ApplicationType
    creator: XlCreator
    parent: object
    connection: object
    text_file_header_row: bool
    text_file_column_data_types: object
    text_file_comma_delimiter: bool
    text_file_consecutive_delimiter: bool
    text_file_decimal_separator: str
    text_file_fixed_column_widths: object
    text_file_other_delimiter: str
    text_file_parse_type: XlTextParsingType
    text_file_platform: XlPlatform
    text_file_prompt_on_refresh: bool
    text_file_semicolon_delimiter: bool
    text_file_space_delimiter: bool
    text_file_start_row: int
    text_file_tab_delimiter: bool
    text_file_text_qualifier: XlTextQualifier
    text_file_thousands_separator: str
    text_file_trailing_minus_numbers: bool
    text_file_visual_layout: XlTextVisualLayoutType
TextConnectionType = TextConnection


class LinkFormat:
    application: ApplicationType
    creator: XlCreator
    parent: object
    auto_update: bool
    locked: bool

    def update(self): pass

LinkFormatType = LinkFormat


class FileExportConverters:

    def __iter__(self, index: int) -> Iterable[FileExportConverterType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, index: int) -> FileExportConverterType: pass

FileExportConvertersType = FileExportConverters


class Model:
    application: ApplicationType
    creator: XlCreator
    parent: object
    model_tables: ModelTablesType
    model_relationships: ModelRelationshipsType

    def refresh(self): pass


    def add_connection(self, connection_to_data_source: WorkbookConnectionType) -> WorkbookConnectionType: pass


    def create_model_workbook_connection(self, model_table: object) -> WorkbookConnectionType: pass

    data_model_connection: WorkbookConnectionType
    name: str

    def initialize(self): pass

    model_measures: ModelMeasuresType
    model_format_general: ModelFormatGeneralType

    def get_model_format_date(self, format_string: object = None) -> ModelFormatDateType: pass


    def get_model_format_decimal_number(self, use_thousand_separator: object = None, decimal_places: object = None) -> ModelFormatDecimalNumberType: pass


    def get_model_format_whole_number(self, use_thousand_separator: object = None) -> ModelFormatWholeNumberType: pass


    def get_model_format_percentage_number(self, use_thousand_separator: object = None, decimal_places: object = None) -> ModelFormatPercentageNumberType: pass


    def get_model_format_scientific_number(self, decimal_places: object = None) -> ModelFormatScientificNumberType: pass


    def get_model_format_currency(self, symbol: object = None, decimal_places: object = None) -> ModelFormatCurrencyType: pass

    model_format_boolean: ModelFormatBooleanType
ModelType = Model


class Hyperlink:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    range: RangeType
    shape: ShapeType
    sub_address: str
    address: str
    type: int

    def add_to_favorites(self): pass


    def delete(self): pass


    def follow(self, new_window: object = None, add_history: object = None, extra_info: object = None, method: object = None, header_info: object = None): pass

    email_subject: str
    screen_tip: str
    text_to_display: str

    def create_new_document(self, filename: str, edit_now: bool, overwrite: bool): pass

HyperlinkType = Hyperlink


class XmlDataBinding:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def refresh(self) -> XlXmlImportResult: pass


    def load_settings(self, url: str): pass


    def clear_settings(self): pass

    source_url: str
XmlDataBindingType = XmlDataBinding


class TableObject:
    application: ApplicationType
    creator: XlCreator
    parent: object
    row_numbers: bool
    fetched_row_overflow: bool
    refresh_style: XlCellInsertionMode
    enable_refresh: bool
    destination: RangeType
    result_range: RangeType

    def delete(self): pass


    def refresh(self) -> bool: pass

    enable_editing: bool
    preserve_column_info: bool
    preserve_formatting: bool
    adjust_column_width: bool
    list_object: ListObjectType
    workbook_connection: WorkbookConnectionType
TableObjectType = TableObject


class Window:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def activate(self) -> object: pass


    def activate_next(self) -> object: pass


    def activate_previous(self) -> object: pass

    active_cell: RangeType
    active_chart: ChartType
    active_pane: PaneType
    active_sheet: WorksheetType
    caption: object

    def close(self, save_changes: object = None, filename: object = None, route_workbook: object = None) -> bool: pass

    display_formulas: bool
    display_gridlines: bool
    display_headings: bool
    display_horizontal_scroll_bar: bool
    display_outline: bool
    display_vertical_scroll_bar: bool
    display_workbook_tabs: bool
    display_zeros: bool
    enable_resize: bool
    freeze_panes: bool
    gridline_color: int
    gridline_color_index: XlColorIndex
    height: float
    index: int

    def large_scroll(self, down: object = None, up: object = None, to_right: object = None, to_left: object = None) -> object: pass

    left: float

    def new_window(self) -> WindowType: pass

    on_window: str
    panes: PanesType

    def print_preview(self, enable_changes: object = None) -> object: pass

    range_selection: RangeType
    scroll_column: int
    scroll_row: int

    def scroll_workbook_tabs(self, sheets: object = None, position: object = None) -> object: pass

    selected_sheets: SheetsType
    selection: object

    def small_scroll(self, down: object = None, up: object = None, to_right: object = None, to_left: object = None) -> object: pass

    split: bool
    split_column: int
    split_horizontal: float
    split_row: int
    split_vertical: float
    tab_ratio: float
    top: float
    type: XlWindowType
    usable_height: float
    usable_width: float
    visible: bool
    visible_range: RangeType
    width: float
    window_number: int
    window_state: XlWindowState
    zoom: object
    view: XlWindowView
    display_right_to_left: bool

    def points_to_screen_pixels_x(self, points: int) -> int: pass


    def points_to_screen_pixels_y(self, points: int) -> int: pass


    def range_from_point(self, x: int, y: int) -> object: pass


    def scroll_into_view(self, left: int, top: int, width: int, height: int, start: object = None): pass

    sheet_views: SheetViewsType
    active_sheet_view: object

    def print_out(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None) -> object: pass

    display_ruler: bool
    auto_filter_date_grouping: bool
    display_whitespace: bool
    hwnd: int
WindowType = Window


class Module:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def activate(self): pass


    def copy(self, before: object = None, after: object = None): pass


    def delete(self): pass

    code_name: str
    index: int

    def move(self, before: object = None, after: object = None): pass

    name: str
    next: object
    on_double_click: str
    on_sheet_activate: str
    on_sheet_deactivate: str
    page_setup: PageSetupType
    previous: object
    protect_contents: bool
    protection_mode: bool

    def select(self, replace: object = None): pass


    def unprotect(self, password: object = None): pass

    visible: XlSheetVisibility
    shapes: ShapesType

    def insert_file(self, filename: object, merge: object = None) -> object: pass


    def save_as(self, filename: str, file_format: object = None, password: object = None, write_res_password: object = None, read_only_recommended: object = None, create_backup: object = None, add_to_mru: object = None, text_codepage: object = None, text_visual_layout: object = None): pass


    def protect(self, password: object = None, drawing_objects: object = None, contents: object = None, scenarios: object = None, user_interface_only: object = None): pass


    def print_out(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None): pass


    def save_as2(self, filename: str, file_format: object = None, password: object = None, write_res_password: object = None, read_only_recommended: object = None, create_backup: object = None, add_to_mru: object = None, text_codepage: object = None, text_visual_layout: object = None): pass

ModuleType = Module


class TickLabels:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def delete(self) -> object: pass

    font: FontType
    name: str
    number_format: str
    number_format_linked: bool
    number_format_local: object
    orientation: XlTickLabelOrientation

    def select(self) -> object: pass

    reading_order: int
    auto_scale_font: object
    depth: int
    offset: int
    alignment: int
    multi_level: bool
    format: ChartFormatType
TickLabelsType = TickLabels


class Corners:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str

    def select(self) -> object: pass

CornersType = Corners


class HPageBreak:
    application: ApplicationType
    creator: XlCreator
    parent: WorksheetType

    def delete(self): pass


    def drag_off(self, direction: XlDirection, region_index: int): pass

    type: XlPageBreak
    extent: XlPageBreakExtent
    location: RangeType
HPageBreakType = HPageBreak


class Worksheet:
    application: ApplicationType
    creator: XlCreator
    parent: SheetsType

    def activate(self): pass


    def copy(self, before: object = None, after: object = None): pass


    def delete(self): pass

    code_name: str
    index: int

    def move(self, before: object = None, after: object = None): pass

    name: str
    next: object
    on_double_click: str
    on_sheet_activate: str
    on_sheet_deactivate: str
    page_setup: PageSetupType
    previous: object

    def print_preview(self, enable_changes: object = None): pass

    protect_contents: bool
    protect_drawing_objects: bool
    protection_mode: bool
    protect_scenarios: bool

    def select(self, replace: object = None): pass


    def unprotect(self, password: object = None): pass

    visible: XlSheetVisibility
    shapes: ShapesType
    transition_exp_eval: bool

    def arcs(self, index: object = None) -> object: pass

    auto_filter_mode: bool

    def set_background_picture(self, filename: str): pass


    def buttons(self, index: object = None) -> object: pass


    def calculate(self): pass

    enable_calculation: bool
    cells: RangeType

    def chart_objects(self, index: object = None) -> object: pass


    def check_boxes(self, index: object = None) -> object: pass


    def check_spelling(self, custom_dictionary: object = None, ignore_uppercase: object = None, always_suggest: object = None, spell_lang: object = None): pass

    circular_reference: RangeType

    def clear_arrows(self): pass

    columns: RangeType
    consolidation_function: XlConsolidationFunction
    consolidation_options: object
    consolidation_sources: object
    display_automatic_page_breaks: bool

    def drawings(self, index: object = None) -> object: pass


    def drawing_objects(self, index: object = None) -> object: pass


    def drop_downs(self, index: object = None) -> object: pass

    enable_auto_filter: bool
    enable_selection: XlEnableSelection
    enable_outlining: bool
    enable_pivot_table: bool

    def evaluate(self, name: object) -> object: pass

    filter_mode: bool

    def reset_all_page_breaks(self): pass


    def group_boxes(self, index: object = None) -> object: pass


    def group_objects(self, index: object = None) -> object: pass


    def labels(self, index: object = None) -> object: pass


    def lines(self, index: object = None) -> object: pass


    def list_boxes(self, index: object = None) -> object: pass

    names: NamesType

    def ole_objects(self, index: object = None) -> object: pass

    on_calculate: str
    on_data: str
    on_entry: str

    def option_buttons(self, index: object = None) -> object: pass

    outline: OutlineType

    def ovals(self, index: object = None) -> object: pass


    def paste(self, destination: object = None, link: object = None): pass


    def pictures(self, index: object = None) -> object: pass


    def pivot_tables(self, index: object = None) -> object: pass


    def pivot_table_wizard(self, source_type: object = None, source_data: object = None, table_destination: object = None, table_name: object = None, row_grand: object = None, column_grand: object = None, save_data: object = None, has_auto_format: object = None, auto_page: object = None, reserved: object = None, background_query: object = None, optimize_cache: object = None, page_field_order: object = None, page_field_wrap_count: object = None, read_data: object = None, connection: object = None) -> PivotTableType: pass


    def get_range(self, cell1: object, cell2: object = None) -> RangeType: pass


    def rectangles(self, index: object = None) -> object: pass

    rows: RangeType

    def scenarios(self, index: object = None) -> object: pass

    scroll_area: str

    def scroll_bars(self, index: object = None) -> object: pass


    def show_all_data(self): pass


    def show_data_form(self): pass


    def spinners(self, index: object = None) -> object: pass

    standard_height: float
    standard_width: float

    def text_boxes(self, index: object = None) -> object: pass

    transition_form_entry: bool
    type: XlSheetType
    used_range: RangeType
    h_page_breaks: HPageBreaksType
    v_page_breaks: VPageBreaksType
    query_tables: QueryTablesType
    display_page_breaks: bool
    comments: CommentsType
    hyperlinks: HyperlinksType

    def clear_circles(self): pass


    def circle_invalid(self): pass

    auto_filter: AutoFilterType
    display_right_to_left: bool
    scripts: None

    def print_out(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None): pass

    tab: TabType
    mail_envelope: None

    def save_as(self, filename: str, file_format: object = None, password: object = None, write_res_password: object = None, read_only_recommended: object = None, create_backup: object = None, add_to_mru: object = None, text_codepage: object = None, text_visual_layout: object = None, local: object = None): pass

    custom_properties: CustomPropertiesType
    smart_tags: SmartTagsType
    protection: ProtectionType

    def paste_special(self, format: object = None, link: object = None, display_as_icon: object = None, icon_file_name: object = None, icon_index: object = None, icon_label: object = None, no_html_formatting: object = None): pass


    def protect(self, password: object = None, drawing_objects: object = None, contents: object = None, scenarios: object = None, user_interface_only: object = None, allow_formatting_cells: object = None, allow_formatting_columns: object = None, allow_formatting_rows: object = None, allow_inserting_columns: object = None, allow_inserting_rows: object = None, allow_inserting_hyperlinks: object = None, allow_deleting_columns: object = None, allow_deleting_rows: object = None, allow_sorting: object = None, allow_filtering: object = None, allow_using_pivot_tables: object = None): pass

    list_objects: ListObjectsType

    def xml_data_query(self, x_path: str, selection_namespaces: object = None, map: object = None) -> RangeType: pass


    def xml_map_query(self, x_path: str, selection_namespaces: object = None, map: object = None) -> RangeType: pass


    def print_out_ex(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None, ignore_print_areas: object = None): pass

    enable_format_conditions_calculation: bool
    sort: SortType

    def export_as_fixed_format(self, type: XlFixedFormatType, filename: object = None, quality: object = None, include_doc_properties: object = None, ignore_print_areas: object = None, from_: object = None, to: object = None, open_after_publish: object = None, fixed_format_ext_class_ptr: object = None): pass

    printed_comment_pages: int

    def export_as_fixed_format2(self, type: XlFixedFormatType, filename: object = None, quality: object = None, include_doc_properties: object = None, ignore_print_areas: object = None, from_: object = None, to: object = None, open_after_publish: object = None, fixed_format_ext_class_ptr: object = None, work_identity: object = None): pass


    def save_as2(self, filename: str, file_format: object = None, password: object = None, write_res_password: object = None, read_only_recommended: object = None, create_backup: object = None, add_to_mru: object = None, text_codepage: object = None, text_visual_layout: object = None, local: object = None): pass

    comments_threaded: CommentsThreadedType
WorksheetType = Worksheet


class Error:
    application: ApplicationType
    creator: XlCreator
    parent: object
    value: bool
    ignore: bool
ErrorType = Error


class FreeformBuilder:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def add_nodes(self, segment_type, editing_type, x1: float, y1: float, x2: object = None, y2: object = None, x3: object = None, y3: object = None): pass


    def convert_to_shape(self) -> ShapeType: pass

FreeformBuilderType = FreeformBuilder


class AllowEditRange:
    title: str
    range: RangeType

    def change_password(self, password: str): pass


    def delete(self): pass


    def unprotect(self, password: object = None): pass

    users: UserAccessListType
AllowEditRangeType = AllowEditRange


class ToolbarButtons:

    def __iter__(self, index: int) -> Iterable[ToolbarButtonType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, button: object = None, before: object = None, on_action: object = None, pushed: object = None, enabled: object = None, status_bar: object = None, help_file: object = None, help_context_id: object = None) -> ToolbarButtonType: pass

    count: int

    def get(self, index: int) -> ToolbarButtonType: pass

ToolbarButtonsType = ToolbarButtons


class DialogFrame:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def copy_picture(self, appearance: XlPictureAppearance = None, format: XlCopyPictureFormat = None) -> object: pass

    height: float
    left: float
    locked: bool
    name: str
    on_action: str

    def select(self, replace: object = None) -> object: pass

    top: float
    width: float
    shape_range: ShapeRangeType
    caption: str

    def get_characters(self, start: object = None, length: object = None) -> CharactersType: pass


    def check_spelling(self, custom_dictionary: object = None, ignore_uppercase: object = None, always_suggest: object = None, spell_lang: object = None) -> object: pass

    locked_text: bool
    text: str
DialogFrameType = DialogFrame


class PivotFields:
    application: ApplicationType
    creator: XlCreator
    parent: PivotTableType
    count: int

    def item(self, index: object) -> object: pass

PivotFieldsType = PivotFields


class ProtectedViewWindow:
    caption: str
    enable_resize: bool
    height: float
    left: float
    top: float
    width: float
    visible: bool
    source_name: str
    source_path: str
    window_state: XlProtectedViewWindowState
    workbook: WorkbookType

    def activate(self): pass


    def close(self) -> bool: pass


    def edit(self, write_res_password: object = None, update_links: object = None) -> WorkbookType: pass

ProtectedViewWindowType = ProtectedViewWindow


class Ranges:

    def __iter__(self, index: int) -> Iterable[RangeType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, index: int) -> RangeType: pass

RangesType = Ranges


class Border:
    application: ApplicationType
    creator: XlCreator
    parent: object
    color: object
    color_index: object
    line_style: object
    weight: object
    theme_color: object
    tint_and_shade: object
BorderType = Border


class WebOptions:
    application: ApplicationType
    creator: XlCreator
    parent: object
    rely_on_css: bool
    organize_in_folder: bool
    use_long_file_names: bool
    download_components: bool
    rely_on_vml: bool
    allow_png: bool
    screen_size: None
    pixels_per_inch: int
    location_of_components: str
    encoding: None
    folder_suffix: str

    def use_default_folder_suffix(self): pass

    target_browser: None
WebOptionsType = WebOptions


class Toolbar:
    application: ApplicationType
    creator: XlCreator
    parent: object
    built_in: bool

    def delete(self): pass

    height: int
    left: int
    name: str
    position: int
    protection: XlToolbarProtection

    def reset(self): pass

    toolbar_buttons: ToolbarButtonsType
    top: int
    visible: bool
    width: int
ToolbarType = Toolbar


class Windows:

    def __iter__(self, index: int) -> Iterable[WindowType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def arrange(self, arrange_style: XlArrangeStyle = None, active_workbook: object = None, sync_horizontal: object = None, sync_vertical: object = None) -> object: pass

    count: int

    def compare_side_by_side_with(self, window_name: object) -> bool: pass


    def break_side_by_side(self) -> bool: pass

    sync_scrolling_side_by_side: bool

    def reset_positions_side_by_side(self): pass


    def get(self, index: int) -> WindowType: pass

WindowsType = Windows


class OLEDBError:
    application: ApplicationType
    creator: XlCreator
    parent: object
    sql_state: str
    error_string: str
    native: int
    number: int
    stage: int
OLEDBErrorType = OLEDBError


class ModelFormatWholeNumber:
    application: ApplicationType
    creator: XlCreator
    parent: object
    use_thousand_separator: bool
ModelFormatWholeNumberType = ModelFormatWholeNumber


class SparkHorizontalAxis:
    application: ApplicationType
    creator: XlCreator
    parent: object
    axis: SparkColorType
    is_date_axis: bool
    right_to_left_plot_order: bool
SparkHorizontalAxisType = SparkHorizontalAxis


class Font:
    application: ApplicationType
    creator: XlCreator
    parent: object
    background: object
    bold: object
    color: object
    color_index: object
    font_style: object
    italic: object
    name: object
    outline_font: object
    shadow: object
    size: object
    strikethrough: object
    subscript: object
    superscript: object
    underline: object
    theme_color: object
    tint_and_shade: object
    theme_font: XlThemeFont
FontType = Font


class CustomProperties:

    def __iter__(self, index: int) -> Iterable[CustomPropertyType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, name: str, value: object) -> CustomPropertyType: pass

    count: int

    def get(self, index: int) -> CustomPropertyType: pass

CustomPropertiesType = CustomProperties


class AutoRecover:
    application: ApplicationType
    creator: XlCreator
    parent: object
    enabled: bool
    time: int
    path: str
AutoRecoverType = AutoRecover


class Application:
    application: ApplicationType
    creator: XlCreator
    parent: ApplicationType
    active_cell: RangeType
    active_chart: ChartType
    active_dialog: DialogSheetType
    active_menu_bar: MenuBarType
    active_printer: str
    active_sheet: WorksheetType
    active_window: WindowType
    active_workbook: WorkbookType
    add_ins: AddInsType
    assistant: None

    def calculate(self): pass

    cells: RangeType
    charts: SheetsType
    columns: RangeType
    command_bars: None
    dde_app_return_code: int

    def dde_execute(self, channel: int, String: str): pass


    def dde_initiate(self, app: str, topic: str) -> int: pass


    def dde_poke(self, channel: int, item: object, data: object): pass


    def dde_request(self, channel: int, item: str) -> object: pass


    def dde_terminate(self, channel: int): pass

    dialog_sheets: SheetsType

    def evaluate(self, name: object) -> object: pass


    def execute_excel4macro(self, String: str) -> object: pass


    def intersect(self, arg1: RangeType, arg2: RangeType, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> RangeType: pass

    menu_bars: MenuBarsType
    modules: ModulesType
    names: NamesType

    def get_range(self, cell1: object, cell2: object = None) -> RangeType: pass

    rows: RangeType

    def run(self, macro: object = None, arg1: object = None, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> object: pass

    selection: RangeType

    def send_keys(self, keys: object, wait: object = None): pass

    sheets: SheetsType

    def get_shortcut_menus(self, index: int) -> MenuType: pass

    this_workbook: WorkbookType
    toolbars: ToolbarsType

    def union(self, arg1: RangeType, arg2: RangeType, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> RangeType: pass

    windows: WindowsType
    workbooks: WorkbooksType
    worksheet_function: WorksheetFunctionType
    worksheets: SheetsType
    excel4intl_macro_sheets: SheetsType
    excel4macro_sheets: SheetsType

    def activate_microsoft_app(self, index: XlMSApplication): pass


    def add_chart_auto_format(self, chart: object, name: str, description: object = None): pass


    def add_custom_list(self, list_array: object, by_row: object = None): pass

    alert_before_overwriting: bool
    alt_startup_path: str
    ask_to_update_links: bool
    enable_animations: bool
    auto_correct: AutoCorrectType
    build: int
    calculate_before_save: bool
    calculation: XlCalculation

    def get_caller(self, index: object = None) -> object: pass

    can_play_sounds: bool
    can_record_sounds: bool
    caption: str
    cell_drag_and_drop: bool

    def centimeters_to_points(self, centimeters: float) -> float: pass


    def check_spelling(self, word: str, custom_dictionary: object = None, ignore_uppercase: object = None) -> bool: pass


    def get_clipboard_formats(self, index: object = None) -> object: pass

    display_clipboard_window: bool
    color_buttons: bool
    command_underlines: XlCommandUnderlines
    constrain_numeric: bool

    def convert_formula(self, formula: object, from_reference_style: XlReferenceStyle, to_reference_style: object = None, to_absolute: object = None, relative_to: object = None) -> object: pass

    copy_objects_with_cells: bool
    cursor: XlMousePointer
    custom_list_count: int
    cut_copy_mode: XlCutCopyMode
    data_entry_mode: int
    default_file_path: str

    def delete_chart_auto_format(self, name: str): pass


    def delete_custom_list(self, list_num: int): pass

    dialogs: DialogsType
    display_alerts: bool
    display_formula_bar: bool
    display_full_screen: bool
    display_note_indicator: bool
    display_comment_indicator: XlCommentDisplayMode
    display_excel4menus: bool
    display_recent_files: bool
    display_scroll_bars: bool
    display_status_bar: bool

    def double_click(self): pass

    edit_directly_in_cell: bool
    enable_auto_complete: bool
    enable_cancel_key: XlEnableCancelKey
    enable_sound: bool
    enable_tip_wizard: bool

    def get_file_converters(self, index1: object = None, index2: object = None) -> object: pass

    file_search: None
    file_find: None
    fixed_decimal: bool
    fixed_decimal_places: int

    def get_custom_list_contents(self, list_num: int) -> object: pass


    def get_custom_list_num(self, list_array: object) -> int: pass


    def get_open_filename(self, file_filter: object = None, filter_index: object = None, title: object = None, button_text: object = None, multi_select: object = None) -> object: pass


    def get_save_as_filename(self, initial_filename: object = None, file_filter: object = None, filter_index: object = None, title: object = None, button_text: object = None) -> object: pass


    def goto(self, reference: object = None, scroll: object = None): pass

    height: float

    def help(self, help_file: object = None, help_context_id: object = None): pass

    ignore_remote_requests: bool

    def inches_to_points(self, inches: float) -> float: pass


    def input_box(self, prompt: str, title: object = None, Default: object = None, left: object = None, top: object = None, help_file: object = None, help_context_id: object = None, type: object = None) -> object: pass

    interactive: bool

    def get_international(self, index: object = None) -> object: pass

    iteration: bool
    large_buttons: bool
    left: float
    library_path: str

    def macro_options(self, macro: object = None, description: object = None, has_menu: object = None, menu_text: object = None, has_shortcut_key: object = None, shortcut_key: object = None, category: object = None, status_bar: object = None, help_context_id: object = None, help_file: object = None): pass


    def mail_logoff(self): pass


    def mail_logon(self, name: object = None, password: object = None, download_new_mail: object = None): pass

    mail_session: object
    mail_system: XlMailSystem
    math_coprocessor_available: bool
    max_change: float
    max_iterations: int
    memory_free: int
    memory_total: int
    memory_used: int
    mouse_available: bool
    move_after_return: bool
    move_after_return_direction: XlDirection
    recent_files: RecentFilesType
    name: str

    def next_letter(self) -> WorkbookType: pass

    network_templates_path: str
    odbc_errors: ODBCErrorsType
    odbc_timeout: int
    on_calculate: str
    on_data: str
    on_double_click: str
    on_entry: str

    def on_key(self, key: str, procedure: object = None): pass


    def on_repeat(self, text: str, procedure: str): pass

    on_sheet_activate: str
    on_sheet_deactivate: str

    def on_time(self, earliest_time: object, procedure: str, latest_time: object = None, schedule: object = None): pass


    def on_undo(self, text: str, procedure: str): pass

    on_window: str
    operating_system: str
    organization_name: str
    path: str
    path_separator: str

    def get_previous_selections(self, index: object = None) -> object: pass

    pivot_table_selection: bool
    prompt_for_summary_info: bool

    def quit(self): pass


    def record_macro(self, basic_code: object = None, xlm_code: object = None): pass

    record_relative: bool
    reference_style: XlReferenceStyle

    def get_registered_functions(self, index1: object = None, index2: object = None) -> object: pass


    def register_xll(self, filename: str) -> bool: pass


    def repeat(self): pass


    def reset_tip_wizard(self): pass

    roll_zoom: bool

    def save(self, filename: object = None): pass


    def save_workspace(self, filename: object = None): pass

    screen_updating: bool

    def set_default_chart(self, format_name: object = None, gallery: object = None): pass

    sheets_in_new_workbook: int
    show_chart_tip_names: bool
    show_chart_tip_values: bool
    standard_font: str
    standard_font_size: float
    startup_path: str
    status_bar: object
    templates_path: str
    show_tool_tips: bool
    top: float
    default_save_format: XlFileFormat
    transition_menu_key: str
    transition_menu_key_action: int
    transition_navig_keys: bool

    def undo(self): pass

    usable_height: float
    usable_width: float
    user_control: bool
    user_name: str
    value: str
    vbe: None
    version: str
    visible: bool

    def volatile(self, Volatile: object = None): pass

    width: float
    windows_for_pens: bool
    window_state: XlWindowState
    ui_language: int
    default_sheet_direction: int
    cursor_movement: int
    control_characters: bool
    enable_events: bool
    display_info_window: bool

    def wait(self, time: object) -> bool: pass

    extend_list: bool
    oledb_errors: OLEDBErrorsType

    def get_phonetic(self, text: object = None) -> str: pass

    com_add_ins: None
    default_web_options: DefaultWebOptionsType
    product_code: str
    user_library_path: str
    auto_percent_entry: bool
    language_settings: None
    answer_wizard: None

    def calculate_full(self): pass


    def find_file(self) -> bool: pass

    calculation_version: int
    show_windows_in_taskbar: bool
    feature_install: None
    ready: bool
    find_format: CellFormatType
    replace_format: CellFormatType
    used_objects: UsedObjectsType
    calculation_state: XlCalculationState
    calculation_interrupt_key: XlCalculationInterruptKey
    watches: WatchesType
    display_function_tool_tips: bool
    automation_security: None

    def get_file_dialog(self, file_dialog_type): pass


    def calculate_full_rebuild(self): pass

    display_paste_options: bool
    display_insert_options: bool
    generate_get_pivot_data: bool
    auto_recover: AutoRecoverType
    hwnd: int
    hinstance: int

    def check_abort(self, keep_abort: object = None): pass

    error_checking_options: ErrorCheckingOptionsType
    auto_format_as_you_type_replace_hyperlinks: bool
    smart_tag_recognizers: SmartTagRecognizersType
    new_workbook: None
    spelling_options: SpellingOptionsType
    speech: SpeechType
    map_paper_size: bool
    show_startup_dialog: bool
    decimal_separator: str
    thousands_separator: str
    use_system_separators: bool
    this_cell: RangeType
    rtd: RTDType
    display_document_action_task_pane: bool

    def display_xml_source_pane(self, xml_map: object = None): pass

    arbitrary_xml_support_available: bool

    def support(self, Object: object, id: int, arg: object = None) -> object: pass

    measurement_unit: int
    show_selection_floaties: bool
    show_menu_floaties: bool
    show_dev_tools: bool
    enable_live_preview: bool
    display_document_information_panel: bool
    always_use_clear_type: bool
    warn_on_function_name_conflict: bool
    formula_bar_height: int
    display_formula_auto_complete: bool
    generate_table_refs: XlGenerateTableRefs
    assistance: None

    def calculate_until_async_queries_done(self): pass

    enable_large_operation_alert: bool
    large_operation_cell_thousand_count: int
    defer_async_queries: bool
    multi_threaded_calculation: MultiThreadedCalculationType

    def share_point_version(self, bstr_url: str) -> int: pass

    active_encryption_session: int
    high_quality_mode_for_graphics: bool
    file_export_converters: FileExportConvertersType
    smart_art_layouts: None
    smart_art_quick_styles: None
    smart_art_colors: None
    add_ins2: AddIns2Type
    print_communication: bool

    def macro_options2(self, macro: object = None, description: object = None, has_menu: object = None, menu_text: object = None, has_shortcut_key: object = None, shortcut_key: object = None, category: object = None, status_bar: object = None, help_context_id: object = None, help_file: object = None, argument_descriptions: object = None): pass

    use_cluster_connector: bool
    cluster_connector: str
    quitting: bool
    protected_view_windows: ProtectedViewWindowsType
    active_protected_view_window: ProtectedViewWindowType
    is_sandboxed: bool
    save_iso8601dates: bool
    hinstance_ptr: object
    file_validation: None
    file_validation_pivot: XlFileValidationPivotMode
    show_quick_analysis: bool
    quick_analysis: QuickAnalysisType
    flash_fill: bool
    enable_macro_animations: bool
    chart_data_point_track: bool
    flash_fill_mode: bool
    merge_instances: bool
    enable_check_file_extensions: bool
    default_pivot_table_layout_options: DefaultPivotTableLayoutOptionsType
ApplicationType = Application


class QueryTables:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def add(self, connection: object, destination: RangeType, sql: object = None) -> QueryTableType: pass


    def item(self, index: object) -> QueryTableType: pass

QueryTablesType = QueryTables


class Filters:

    def __iter__(self, index: int) -> Iterable[FilterType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, index: int) -> FilterType: pass

FiltersType = Filters


class CustomView:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    print_settings: bool
    row_col_settings: bool

    def show(self): pass


    def delete(self): pass

CustomViewType = CustomView


class SlicerCacheLevels:

    def __iter__(self, level: int) -> Iterable[SlicerCacheLevelType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, level: int = None) -> SlicerCacheLevelType: pass

SlicerCacheLevelsType = SlicerCacheLevels


class SlicerCache:
    application: ApplicationType
    creator: XlCreator
    parent: object
    index: int
    olap: bool
    source_type: XlPivotTableSourceType
    workbook_connection: WorkbookConnectionType
    slicers: SlicersType
    pivot_tables: SlicerPivotTablesType
    slicer_cache_levels: SlicerCacheLevelsType
    name: str
    visible_slicer_items: SlicerItemsType
    visible_slicer_items_list: object
    slicer_items: SlicerItemsType
    cross_filter_type: XlSlicerCrossFilterType
    sort_items: XlSlicerSort
    source_name: str
    sort_using_custom_lists: bool
    show_all_items: bool

    def clear_manual_filter(self): pass


    def delete(self): pass

    timeline_state: TimelineStateType

    def clear_all_filters(self): pass

    slicer_cache_type: XlSlicerCacheType
    filter_cleared: bool
    list: bool
    require_manual_update: bool
    list_object: ListObjectType

    def clear_date_filter(self): pass

SlicerCacheType = SlicerCache


class CalculatedMembers:

    def __iter__(self, index: int) -> Iterable[CalculatedMemberType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def add(self, name: str, formula: str, solve_order: object = None, type: object = None) -> CalculatedMemberType: pass


    def add2(self, name: str, formula: object, solve_order: object = None, type: object = None, dynamic: object = None, display_folder: object = None, hierarchize_distinct: object = None) -> CalculatedMemberType: pass


    def add_calculated_member(self, name: str, formula: object, solve_order: object = None, type: object = None, display_folder: object = None, measure_group: object = None, parent_hierarchy: object = None, parent_member: object = None, number_format: object = None) -> CalculatedMemberType: pass


    def get(self, index: int) -> CalculatedMemberType: pass

CalculatedMembersType = CalculatedMembers


class Toolbars:

    def __iter__(self, index: int) -> Iterable[ToolbarType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, name: object = None) -> ToolbarType: pass

    count: int

    def get(self, index: int) -> ToolbarType: pass

ToolbarsType = Toolbars


class ServerViewableItems:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def add(self, obj: object) -> object: pass


    def delete(self, index: object): pass


    def delete_all(self): pass


    def item(self, index: object) -> object: pass

ServerViewableItemsType = ServerViewableItems


class DiagramNode:

    def add_node(self, pos = None, node_type = None) -> DiagramNodeType: pass


    def delete(self): pass


    def move_node(self, p_target_node: DiagramNodeType, pos): pass


    def replace_node(self, p_target_node: DiagramNodeType): pass


    def swap_node(self, p_target_node: DiagramNodeType, swap_children: bool = None): pass


    def clone_node(self, copy_children: bool, p_target_node: DiagramNodeType, pos = None) -> DiagramNodeType: pass


    def transfer_children(self, p_receiving_node: DiagramNodeType): pass


    def next_node(self) -> DiagramNodeType: pass


    def prev_node(self) -> DiagramNodeType: pass

    parent: object
    children: DiagramNodeChildrenType
    shape: ShapeType
    root: DiagramNodeType
    diagram: None
    layout: None
    text_shape: ShapeType
DiagramNodeType = DiagramNode


class ShapeRange:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: object) -> ShapeType: pass


    def align(self, align_cmd, relative_to): pass


    def apply(self): pass


    def delete(self): pass


    def distribute(self, distribute_cmd, relative_to): pass


    def duplicate(self) -> ShapeRangeType: pass


    def flip(self, flip_cmd): pass


    def increment_left(self, increment: float): pass


    def increment_rotation(self, increment: float): pass


    def increment_top(self, increment: float): pass


    def group(self) -> ShapeType: pass


    def pick_up(self): pass


    def reroute_connections(self): pass


    def regroup(self) -> ShapeType: pass


    def scale_height(self, factor: float, relative_to_original_size, scale: object = None): pass


    def scale_width(self, factor: float, relative_to_original_size, scale: object = None): pass


    def select(self, replace: object = None): pass


    def set_shapes_default_properties(self): pass


    def ungroup(self) -> ShapeRangeType: pass


    def z_order(self, z_order_cmd): pass

    adjustments: AdjustmentsType
    text_frame: TextFrameType
    auto_shape_type: None
    callout: CalloutFormatType
    connection_site_count: int
    connector: None
    connector_format: ConnectorFormatType
    fill: FillFormatType
    group_items: GroupShapesType
    height: float
    horizontal_flip: None
    left: float
    line: LineFormatType
    lock_aspect_ratio: None
    name: str
    nodes: ShapeNodesType
    rotation: float
    picture_format: PictureFormatType
    shadow: ShadowFormatType
    text_effect: TextEffectFormatType
    three_d: ThreeDFormatType
    top: float
    type: None
    vertical_flip: None
    vertices: object
    visible: None
    width: float
    z_order_position: int
    black_white_mode: None
    alternative_text: str
    diagram_node: DiagramNodeType
    has_diagram_node: None
    diagram: DiagramType
    has_diagram: None
    child: None
    parent_group: ShapeType
    canvas_items: None
    id: int

    def canvas_crop_left(self, increment: float): pass


    def canvas_crop_top(self, increment: float): pass


    def canvas_crop_right(self, increment: float): pass


    def canvas_crop_bottom(self, increment: float): pass

    chart: ChartType
    has_chart: None
    text_frame2: TextFrame2Type
    shape_style: None
    background_style: None
    soft_edge: None
    glow: None
    reflection: None
    title: str
    graphic_style: None
    model3d: Model3DFormatType
ShapeRangeType = ShapeRange


class Phonetic:
    application: ApplicationType
    creator: XlCreator
    parent: object
    visible: bool
    character_type: int
    alignment: int
    font: FontType
    text: str
PhoneticType = Phonetic


class VPageBreak:
    application: ApplicationType
    creator: XlCreator
    parent: WorksheetType

    def delete(self): pass


    def drag_off(self, direction: XlDirection, region_index: int): pass

    type: XlPageBreak
    extent: XlPageBreakExtent
    location: RangeType
VPageBreakType = VPageBreak


class PivotItemList:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: object) -> PivotItemType: pass

PivotItemListType = PivotItemList


class TextEffectFormat:
    parent: object

    def toggle_vertical_text(self): pass

    alignment: None
    font_bold: None
    font_italic: None
    font_name: str
    font_size: float
    kerned_pairs: None
    normalized_height: None
    preset_shape: None
    preset_text_effect: None
    rotated_chars: None
    text: str
    tracking: float
TextEffectFormatType = TextEffectFormat


class SlicerCacheLevel:
    application: ApplicationType
    creator: XlCreator
    parent: object
    slicer_items: SlicerItemsType
    count: int
    ordinal: int
    name: str
    cross_filter_type: XlSlicerCrossFilterType
    sort_items: XlSlicerSort
    visible_slicer_items_list: object
SlicerCacheLevelType = SlicerCacheLevel


class FormatColor:
    application: ApplicationType
    creator: XlCreator
    parent: object
    color: object
    color_index: XlColorIndex
    theme_color: object
    tint_and_shade: object
FormatColorType = FormatColor


class DiagramNodeChildren:

    def item(self, index: object) -> DiagramNodeType: pass


    def add_node(self, index: int = None, node_type = None) -> DiagramNodeType: pass


    def select_all(self): pass

    parent: object
    count: int
    first_child: DiagramNodeType
    last_child: DiagramNodeType
DiagramNodeChildrenType = DiagramNodeChildren


class ListRows:

    def __iter__(self, index: int) -> Iterable[ListRowType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, position: object = None) -> ListRowType: pass

    count: int

    def add_ex(self, position: object = None, always_insert: object = None) -> ListRowType: pass


    def get(self, index: int) -> ListRowType: pass

ListRowsType = ListRows


class SmartTags:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, smart_tag_type: str) -> SmartTagType: pass

    count: int
SmartTagsType = SmartTags


class UpBars:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str

    def select(self) -> object: pass

    border: BorderType

    def delete(self) -> object: pass

    interior: InteriorType
    fill: ChartFillFormatType
    format: ChartFormatType

    def set_property(self, id: str, value: object): pass


    def get_property(self, id: str) -> object: pass

UpBarsType = UpBars


class Range:

    def __iter__(self, row_index: int, column_index: int) -> Iterable[RangeType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: WorksheetType

    def activate(self) -> object: pass

    add_indent: object

    def get_address(self, row_absolute: object = None, column_absolute: object = None, reference_style: XlReferenceStyle = None, external: object = None, relative_to: object = None) -> str: pass


    def get_address_local(self, row_absolute: object = None, column_absolute: object = None, reference_style: XlReferenceStyle = None, external: object = None, relative_to: object = None) -> str: pass


    def advanced_filter(self, action: XlFilterAction, criteria_range: object = None, copy_to_range: object = None, unique: object = None) -> object: pass


    def apply_names(self, names: object = None, ignore_relative_absolute: object = None, use_row_column_names: object = None, omit_column: object = None, omit_row: object = None, order: XlApplyNamesOrder = None, append_last: object = None) -> object: pass


    def apply_outline_styles(self) -> object: pass

    areas: AreasType

    def auto_complete(self, String: str) -> str: pass


    def auto_fill(self, destination: RangeType, type: XlAutoFillType = None) -> object: pass


    def auto_filter(self, field: object = None, criteria1: object = None, Operator: XlAutoFilterOperator = None, criteria2: object = None, visible_drop_down: object = None) -> object: pass


    def auto_fit(self) -> object: pass


    def auto_format(self, format: XlRangeAutoFormat = None, number: object = None, font: object = None, alignment: object = None, border: object = None, pattern: object = None, width: object = None) -> object: pass


    def auto_outline(self) -> object: pass


    def border_around(self, line_style: object = None, weight: XlBorderWeight = None, color_index: XlColorIndex = None, color: object = None) -> object: pass

    borders: BordersType

    def calculate(self) -> object: pass

    cells: RangeType

    def get_characters(self, start: object = None, length: object = None) -> CharactersType: pass


    def check_spelling(self, custom_dictionary: object = None, ignore_uppercase: object = None, always_suggest: object = None, spell_lang: object = None) -> object: pass


    def clear(self) -> object: pass


    def clear_contents(self) -> object: pass


    def clear_formats(self) -> object: pass


    def clear_notes(self) -> object: pass


    def clear_outline(self) -> object: pass

    column: int

    def column_differences(self, comparison: object) -> RangeType: pass

    columns: RangeType
    column_width: object

    def consolidate(self, sources: object = None, function: object = None, top_row: object = None, left_column: object = None, create_links: object = None) -> object: pass


    def copy(self, destination: object = None) -> object: pass


    def copy_from_recordset(self, data: object, max_rows: object = None, max_columns: object = None) -> int: pass


    def copy_picture(self, appearance: XlPictureAppearance = None, format: XlCopyPictureFormat = None) -> object: pass

    count: int

    def create_names(self, top: object = None, left: object = None, bottom: object = None, right: object = None) -> object: pass


    def create_publisher(self, edition: object = None, appearance: XlPictureAppearance = None, contains_pict: object = None, contains_biff: object = None, contains_rtf: object = None, contains_valu: object = None) -> object: pass

    current_array: RangeType
    current_region: RangeType

    def cut(self, destination: object = None) -> object: pass


    def data_series(self, rowcol: object = None, type: XlDataSeriesType = None, date: XlDataSeriesDate = None, step: object = None, stop: object = None, trend: object = None) -> object: pass


    def delete(self, shift: object = None) -> object: pass

    dependents: RangeType

    def dialog_box(self) -> object: pass

    direct_dependents: RangeType
    direct_precedents: RangeType

    def edition_options(self, type: XlEditionType, option: XlEditionOptionsOption, name: object = None, reference: object = None, appearance: XlPictureAppearance = None, chart_size: XlPictureAppearance = None, format: object = None) -> object: pass


    def get_end(self, direction: XlDirection) -> RangeType: pass

    entire_column: RangeType
    entire_row: RangeType

    def fill_down(self) -> object: pass


    def fill_left(self) -> object: pass


    def fill_right(self) -> object: pass


    def fill_up(self) -> object: pass


    def find(self, what: object, after: object = None, look_in: object = None, look_at: object = None, search_order: object = None, search_direction: XlSearchDirection = None, match_case: object = None, match_byte: object = None, search_format: object = None) -> RangeType: pass


    def find_next(self, after: object = None) -> RangeType: pass


    def find_previous(self, after: object = None) -> RangeType: pass

    font: FontType
    formula: object
    formula_array: object
    formula_label: XlFormulaLabel
    formula_hidden: object
    formula_local: object
    formula_r1c1: object
    formula_r1c1local: object

    def function_wizard(self) -> object: pass


    def goal_seek(self, goal: object, changing_cell: RangeType) -> bool: pass


    def group(self, start: object = None, end: object = None, by: object = None, periods: object = None) -> object: pass

    has_array: object
    has_formula: object
    height: object
    hidden: object
    horizontal_alignment: object
    indent_level: object

    def insert_indent(self, insert_amount: int): pass


    def insert(self, shift: object = None, copy_origin: object = None) -> object: pass

    interior: InteriorType

    def justify(self) -> object: pass

    left: object
    list_header_rows: int

    def list_names(self) -> object: pass

    location_in_table: XlLocationInTable
    locked: object

    def merge(self, across: object = None): pass


    def un_merge(self): pass

    merge_area: RangeType
    merge_cells: object
    name: object

    def navigate_arrow(self, toward_precedent: object = None, arrow_number: object = None, link_number: object = None) -> object: pass

    next: RangeType

    def note_text(self, text: object = None, start: object = None, length: object = None) -> str: pass

    number_format: object
    number_format_local: object

    def get_offset(self, row_offset: object = None, column_offset: object = None) -> RangeType: pass

    orientation: object
    outline_level: object
    page_break: int

    def parse(self, parse_line: object = None, destination: object = None) -> object: pass

    pivot_field: PivotFieldType
    pivot_item: PivotItemType
    pivot_table: PivotTableType
    precedents: RangeType
    prefix_character: object
    previous: RangeType

    def print_preview(self, enable_changes: object = None) -> object: pass

    query_table: QueryTableType

    def get_range(self, cell1: object, cell2: object = None) -> RangeType: pass


    def remove_subtotal(self) -> object: pass


    def replace(self, what: object, replacement: object, look_at: object = None, search_order: object = None, match_case: object = None, match_byte: object = None, search_format: object = None, replace_format: object = None) -> bool: pass


    def get_resize(self, row_size: object = None, column_size: object = None) -> RangeType: pass

    row: int

    def row_differences(self, comparison: object) -> RangeType: pass

    row_height: object
    rows: RangeType

    def run(self, arg1: object = None, arg2: object = None, arg3: object = None, arg4: object = None, arg5: object = None, arg6: object = None, arg7: object = None, arg8: object = None, arg9: object = None, arg10: object = None, arg11: object = None, arg12: object = None, arg13: object = None, arg14: object = None, arg15: object = None, arg16: object = None, arg17: object = None, arg18: object = None, arg19: object = None, arg20: object = None, arg21: object = None, arg22: object = None, arg23: object = None, arg24: object = None, arg25: object = None, arg26: object = None, arg27: object = None, arg28: object = None, arg29: object = None, arg30: object = None) -> object: pass


    def select(self) -> object: pass


    def show(self) -> object: pass


    def show_dependents(self, remove: object = None) -> object: pass

    show_detail: object

    def show_errors(self) -> object: pass


    def show_precedents(self, remove: object = None) -> object: pass

    shrink_to_fit: object

    def sort(self, key1: object = None, order1: XlSortOrder = None, key2: object = None, type: object = None, order2: XlSortOrder = None, key3: object = None, order3: XlSortOrder = None, header: XlYesNoGuess = None, order_custom: object = None, match_case: object = None, orientation: XlSortOrientation = None, sort_method: XlSortMethod = None, data_option1: XlSortDataOption = None, data_option2: XlSortDataOption = None, data_option3: XlSortDataOption = None) -> object: pass


    def sort_special(self, sort_method: XlSortMethod = None, key1: object = None, order1: XlSortOrder = None, type: object = None, key2: object = None, order2: XlSortOrder = None, key3: object = None, order3: XlSortOrder = None, header: XlYesNoGuess = None, order_custom: object = None, match_case: object = None, orientation: XlSortOrientation = None, data_option1: XlSortDataOption = None, data_option2: XlSortDataOption = None, data_option3: XlSortDataOption = None) -> object: pass

    sound_note: SoundNoteType

    def special_cells(self, type: XlCellType, value: object = None) -> RangeType: pass

    style: object

    def subscribe_to(self, edition: str, format: XlSubscribeToFormat = None) -> object: pass


    def subtotal(self, group_by: int, function: XlConsolidationFunction, total_list: object, replace: object = None, page_breaks: object = None, summary_below_data: XlSummaryRow = None) -> object: pass

    summary: object

    def table(self, row_input: object = None, column_input: object = None) -> object: pass

    text: object

    def text_to_columns(self, destination: object = None, data_type: XlTextParsingType = None, text_qualifier: XlTextQualifier = None, consecutive_delimiter: object = None, tab: object = None, semicolon: object = None, comma: object = None, space: object = None, other: object = None, other_char: object = None, field_info: object = None, decimal_separator: object = None, thousands_separator: object = None, trailing_minus_numbers: object = None) -> object: pass

    top: object

    def ungroup(self) -> object: pass

    use_standard_height: object
    use_standard_width: object
    validation: ValidationType

    def get_value(self, range_value_data_type: object = None) -> object: pass


    def set_value(self, range_value_data_type: object = None, _param2: object = None): pass

    vertical_alignment: object
    width: object
    worksheet: WorksheetType
    wrap_text: object

    def add_comment(self, text: object = None) -> CommentType: pass

    comment: CommentType

    def clear_comments(self): pass

    phonetic: PhoneticType
    format_conditions: FormatConditionsType
    reading_order: int
    hyperlinks: HyperlinksType
    phonetics: PhoneticsType

    def set_phonetic(self): pass

    id: str

    def print_out(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None) -> object: pass

    pivot_cell: PivotCellType

    def dirty(self): pass

    errors: ErrorsType
    smart_tags: SmartTagsType

    def speak(self, speak_direction: object = None, speak_formulas: object = None): pass


    def paste_special(self, paste: XlPasteType = None, operation: XlPasteSpecialOperation = None, skip_blanks: object = None, transpose: object = None) -> object: pass

    allow_edit: bool
    list_object: ListObjectType
    x_path: XPathType
    server_actions: ActionsType

    def remove_duplicates(self, columns: object = None, header: XlYesNoGuess = None): pass


    def print_out_ex(self, from_: object = None, to: object = None, copies: object = None, preview: object = None, active_printer: object = None, print_to_file: object = None, collate: object = None, pr_to_file_name: object = None) -> object: pass

    mdx: str

    def export_as_fixed_format(self, type: XlFixedFormatType, filename: object = None, quality: object = None, include_doc_properties: object = None, ignore_print_areas: object = None, from_: object = None, to: object = None, open_after_publish: object = None, fixed_format_ext_class_ptr: object = None): pass

    count_large: object

    def calculate_row_major_order(self) -> object: pass

    sparkline_groups: SparklineGroupsType

    def clear_hyperlinks(self): pass

    display_format: DisplayFormatType

    def border_around2(self, line_style: object = None, weight: XlBorderWeight = None, color_index: XlColorIndex = None, color: object = None, theme_color: object = None) -> object: pass


    def allocate_changes(self): pass


    def discard_changes(self): pass


    def flash_fill(self): pass


    def export_as_fixed_format2(self, type: XlFixedFormatType, filename: object = None, quality: object = None, include_doc_properties: object = None, ignore_print_areas: object = None, from_: object = None, to: object = None, open_after_publish: object = None, fixed_format_ext_class_ptr: object = None, work_identity: object = None): pass


    def add_comment_threaded(self, text: object = None) -> CommentThreadedType: pass

    comment_threaded: CommentThreadedType
    value: object

    def get(self, row_index: int, column_index: int) -> RangeType: pass


    def set(self, row_index: int, column_index: int, _param3: RangeType): pass

RangeType = Range


class PublishObject:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def delete(self): pass


    def publish(self, create: object = None): pass

    div_id: str
    sheet: str
    source_type: XlSourceType
    source: str
    html_type: XlHtmlType
    title: str
    filename: str
    auto_republish: bool
PublishObjectType = PublishObject


class PivotFormulas:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: object) -> PivotFormulaType: pass


    def add(self, formula: str, use_standard_formula: object = None) -> PivotFormulaType: pass

PivotFormulasType = PivotFormulas


class PivotCaches:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: object) -> PivotCacheType: pass


    def add(self, source_type: XlPivotTableSourceType, source_data: object = None) -> PivotCacheType: pass


    def create(self, source_type: XlPivotTableSourceType, source_data: object = None, version: object = None) -> PivotCacheType: pass

PivotCachesType = PivotCaches


class SparklineGroups:

    def __iter__(self, index: int) -> Iterable[SparklineGroupType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, type: XlSparkType, source_data: str) -> SparklineGroupType: pass

    count: int

    def clear(self): pass


    def clear_groups(self): pass


    def group(self, location: RangeType): pass


    def ungroup(self): pass


    def get(self, index: int) -> SparklineGroupType: pass

SparklineGroupsType = SparklineGroups


class PivotFormula:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def delete(self): pass

    formula: str
    value: str
    index: int
    standard_formula: str
PivotFormulaType = PivotFormula


class ThreeDFormat:
    parent: object

    def increment_rotation_x(self, increment: float): pass


    def increment_rotation_y(self, increment: float): pass


    def reset_rotation(self): pass


    def set_three_d_format(self, preset_three_d_format): pass


    def set_extrusion_direction(self, preset_extrusion_direction): pass

    depth: float
    extrusion_color: ColorFormatType
    extrusion_color_type: None
    perspective: None
    preset_extrusion_direction: None
    preset_lighting_direction: None
    preset_lighting_softness: None
    preset_material: None
    preset_three_d_format: None
    rotation_x: float
    rotation_y: float
    visible: None

    def set_preset_camera(self, preset_camera): pass


    def increment_rotation_z(self, increment: float): pass


    def increment_rotation_horizontal(self, increment: float): pass


    def increment_rotation_vertical(self, increment: float): pass

    preset_lighting: None
    z: float
    bevel_top_type: None
    bevel_top_inset: float
    bevel_top_depth: float
    bevel_bottom_type: None
    bevel_bottom_inset: float
    bevel_bottom_depth: float
    rotation_z: float
    contour_width: float
    contour_color: ColorFormatType
    field_of_view: float
    project_text: None
    light_angle: float
    preset_camera_: None
ThreeDFormatType = ThreeDFormat


class Areas:

    def __iter__(self, index: int) -> Iterable[RangeType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, index: int) -> RangeType: pass

AreasType = Areas


class Graphic:
    application: ApplicationType
    creator: XlCreator
    parent: object
    brightness: float
    color_type: None
    contrast: float
    crop_bottom: float
    crop_left: float
    crop_right: float
    crop_top: float
    filename: str
    height: float
    lock_aspect_ratio: None
    width: float
GraphicType = Graphic


class RTD:
    throttle_interval: int

    def refresh_data(self): pass


    def restart_servers(self): pass

RTDType = RTD


class VPageBreaks:

    def __iter__(self, index: int) -> Iterable[VPageBreakType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def add(self, before: object) -> VPageBreakType: pass


    def get(self, index: int) -> VPageBreakType: pass

VPageBreaksType = VPageBreaks


class PivotFilters:

    def __iter__(self, index: int) -> Iterable[PivotFilterType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def add(self, type: XlPivotFilterType, data_field: object = None, value1: object = None, value2: object = None, order: object = None, name: object = None, description: object = None, member_property_field: object = None) -> PivotFilterType: pass


    def add2(self, type: XlPivotFilterType, data_field: object = None, value1: object = None, value2: object = None, order: object = None, name: object = None, description: object = None, member_property_field: object = None, whole_day_filter: object = None) -> PivotFilterType: pass


    def get(self, index: int) -> PivotFilterType: pass

PivotFiltersType = PivotFilters


class QueryTable:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    field_names: bool
    row_numbers: bool
    fill_adjacent_formulas: bool
    has_auto_format: bool
    refresh_on_file_open: bool
    refreshing: bool
    fetched_row_overflow: bool
    background_query: bool

    def cancel_refresh(self): pass

    refresh_style: XlCellInsertionMode
    enable_refresh: bool
    save_password: bool
    destination: RangeType
    connection: object
    sql: object
    post_text: str
    result_range: RangeType

    def delete(self): pass


    def refresh(self, background_query: object = None) -> bool: pass

    parameters: ParametersType
    recordset: object
    save_data: bool
    tables_only_from_html: bool
    enable_editing: bool
    text_file_platform: int
    text_file_start_row: int
    text_file_parse_type: XlTextParsingType
    text_file_text_qualifier: XlTextQualifier
    text_file_consecutive_delimiter: bool
    text_file_tab_delimiter: bool
    text_file_semicolon_delimiter: bool
    text_file_comma_delimiter: bool
    text_file_space_delimiter: bool
    text_file_other_delimiter: str
    text_file_column_data_types: object
    text_file_fixed_column_widths: object
    preserve_column_info: bool
    preserve_formatting: bool
    adjust_column_width: bool
    command_text: object
    command_type: XlCmdType
    text_file_prompt_on_refresh: bool
    query_type: XlQueryType
    maintain_connection: bool
    text_file_decimal_separator: str
    text_file_thousands_separator: str
    refresh_period: int

    def reset_timer(self): pass

    web_selection_type: XlWebSelectionType
    web_formatting: XlWebFormatting
    web_tables: str
    web_pre_formatted_text_to_columns: bool
    web_single_block_text_import: bool
    web_disable_date_recognition: bool
    web_consecutive_delimiters_as_one: bool
    web_disable_redirections: bool
    edit_web_page: object
    source_connection_file: str
    source_data_file: str
    robust_connect: XlRobustConnect
    text_file_trailing_minus_numbers: bool

    def save_as_odc(self, odc_file_name: str, description: object = None, keywords: object = None): pass

    list_object: ListObjectType
    text_file_visual_layout: XlTextVisualLayoutType
    workbook_connection: WorkbookConnectionType
    sort: SortType
QueryTableType = QueryTable


class Pane:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def activate(self) -> bool: pass

    index: int

    def large_scroll(self, down: object = None, up: object = None, to_right: object = None, to_left: object = None) -> object: pass

    scroll_column: int
    scroll_row: int

    def small_scroll(self, down: object = None, up: object = None, to_right: object = None, to_left: object = None) -> object: pass

    visible_range: RangeType

    def scroll_into_view(self, left: int, top: int, width: int, height: int, start: object = None): pass


    def points_to_screen_pixels_x(self, points: int) -> int: pass


    def points_to_screen_pixels_y(self, points: int) -> int: pass

PaneType = Pane


class UserAccess:
    name: str
    allow_edit: bool

    def delete(self): pass

UserAccessType = UserAccess


class RecentFile:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    path: str
    index: int

    def open(self) -> WorkbookType: pass


    def delete(self): pass

RecentFileType = RecentFile


class PivotLayout:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def get_column_fields(self, index: object = None) -> object: pass


    def get_data_fields(self, index: object = None) -> object: pass


    def get_page_fields(self, index: object = None) -> object: pass


    def get_row_fields(self, index: object = None) -> object: pass


    def get_hidden_fields(self, index: object = None) -> object: pass


    def get_visible_fields(self, index: object = None) -> object: pass


    def get_pivot_fields(self, index: object = None) -> object: pass

    cube_fields: CubeFieldsType
    pivot_cache: PivotCacheType
    pivot_table: PivotTableType
    inner_detail: str

    def add_fields(self, row_fields: object = None, column_fields: object = None, page_fields: object = None, append_field: object = None): pass

PivotLayoutType = PivotLayout


class TextFrame2:
    parent: object
    margin_bottom: float
    margin_left: float
    margin_right: float
    margin_top: float
    orientation: None
    horizontal_anchor: None
    vertical_anchor: None
    path_format: None
    warp_format: None
    word_artformat: None
    word_wrap: None
    auto_size: None
    three_d: ThreeDFormatType
    has_text: None
    text_range: None
    column: None
    ruler: None

    def delete_text(self): pass

    no_text_rotation: None
TextFrame2Type = TextFrame2


class SparkVerticalAxis:
    application: ApplicationType
    creator: XlCreator
    parent: object
    min_scale_type: XlSparkScale
    custom_min_scale_value: object
    max_scale_type: XlSparkScale
    custom_max_scale_value: object
SparkVerticalAxisType = SparkVerticalAxis


class CalculatedMember:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    formula: str
    source_name: str
    solve_order: int
    is_valid: bool

    def delete(self): pass

    type: XlCalculatedMemberType
    dynamic: bool
    display_folder: str
    hierarchize_distinct: bool
    flatten_hierarchies: bool
    measure_group: str
    parent_hierarchy: str
    parent_member: str
    number_format: XlCalcMemNumberFormatType
CalculatedMemberType = CalculatedMember


class ModelRelationship:
    application: ApplicationType
    creator: XlCreator
    parent: object
    foreign_key_table: ModelTableType
    foreign_key_column: ModelTableColumnType
    primary_key_table: ModelTableType
    primary_key_column: ModelTableColumnType
    active: bool

    def delete(self): pass

ModelRelationshipType = ModelRelationship


class Walls:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str

    def select(self) -> object: pass

    border: BorderType

    def clear_formats(self) -> object: pass

    interior: InteriorType
    fill: ChartFillFormatType
    picture_type: object

    def paste(self): pass

    picture_unit: object
    thickness: int
    format: ChartFormatType
WallsType = Walls


class DataFeedConnection:
    application: ApplicationType
    creator: XlCreator
    parent: object
    always_use_connection_file: bool
    command_text: object
    command_type: XlCmdType
    connection: object
    enable_refresh: bool
    refresh_date: None
    refreshing: bool
    refresh_on_file_open: bool
    refresh_period: int
    save_password: bool
    server_credentials_method: XlCredentialsMethod
    source_connection_file: str
    source_data_file: str

    def cancel_refresh(self): pass


    def refresh(self): pass


    def save_as_odc(self, odc_file_name: str, description: object = None, keywords: object = None): pass

DataFeedConnectionType = DataFeedConnection


class AllowEditRanges:

    def __iter__(self, index: int) -> Iterable[AllowEditRangeType]: pass

    count: int

    def add(self, title: str, range: RangeType, password: object = None) -> AllowEditRangeType: pass


    def get(self, index: int) -> AllowEditRangeType: pass

AllowEditRangesType = AllowEditRanges


class ShapeNode:
    parent: object
    editing_type: None
    points: object
    segment_type: None
ShapeNodeType = ShapeNode


class Protection:
    allow_formatting_cells: bool
    allow_formatting_columns: bool
    allow_formatting_rows: bool
    allow_inserting_columns: bool
    allow_inserting_rows: bool
    allow_inserting_hyperlinks: bool
    allow_deleting_columns: bool
    allow_deleting_rows: bool
    allow_sorting: bool
    allow_filtering: bool
    allow_using_pivot_tables: bool
    allow_edit_ranges: AllowEditRangesType
ProtectionType = Protection


class Mailer:
    application: ApplicationType
    creator: XlCreator
    parent: object
    bcc_recipients: object
    cc_recipients: object
    enclosures: object
    received: bool
    send_date_time: None
    sender: str
    subject: str
    to_recipients: object
    which_address: object
MailerType = Mailer


class ModelFormatPercentageNumber:
    application: ApplicationType
    creator: XlCreator
    parent: object
    use_thousand_separator: bool
    decimal_places: int
ModelFormatPercentageNumberType = ModelFormatPercentageNumber


class RecentFiles:

    def __iter__(self, index: int) -> Iterable[RecentFileType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    maximum: int
    count: int

    def add(self, name: str) -> RecentFileType: pass


    def get(self, index: int) -> RecentFileType: pass

RecentFilesType = RecentFiles


class ErrorCheckingOptions:
    application: ApplicationType
    creator: XlCreator
    parent: object
    background_checking: bool
    indicator_color_index: XlColorIndex
    evaluate_to_error: bool
    text_date: bool
    number_as_text: bool
    inconsistent_formula: bool
    omitted_cells: bool
    unlocked_formula_cells: bool
    empty_cell_references: bool
    list_data_validation: bool
    inconsistent_table_formula: bool
ErrorCheckingOptionsType = ErrorCheckingOptions


class Dialogs:

    def __iter__(self, index: XlBuiltInDialog) -> Iterable[DialogType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, index: XlBuiltInDialog) -> DialogType: pass

DialogsType = Dialogs


class Comments:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: int) -> CommentType: pass

CommentsType = Comments


class Diagram:
    application: ApplicationType
    creator: XlCreator
    parent: object
    nodes: DiagramNodesType
    type: None
    auto_layout: None
    reverse: None
    auto_format: None

    def convert(self, type): pass


    def fit_text(self): pass

DiagramType = Diagram


class ModelFormatDecimalNumber:
    application: ApplicationType
    creator: XlCreator
    parent: object
    use_thousand_separator: bool
    decimal_places: int
ModelFormatDecimalNumberType = ModelFormatDecimalNumber


class ModelTableColumns:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: object) -> ModelTableColumnType: pass

ModelTableColumnsType = ModelTableColumns


class Validation:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, type: XlDVType, alert_style: object = None, Operator: object = None, formula1: object = None, formula2: object = None): pass

    alert_style: int
    ignore_blank: bool
    ime_mode: int
    in_cell_dropdown: bool

    def delete(self): pass

    error_message: str
    error_title: str
    input_message: str
    input_title: str
    formula1: str
    formula2: str

    def modify(self, type: object = None, alert_style: object = None, Operator: object = None, formula1: object = None, formula2: object = None): pass

    operator: int
    show_error: bool
    show_input: bool
    type: int
    value: bool
ValidationType = Validation


class XPath:
    application: ApplicationType
    creator: XlCreator
    parent: object
    value: str
    map: XmlMapType

    def set_value(self, map: XmlMapType, x_path: str, selection_namespace: object = None, repeating: object = None): pass


    def clear(self): pass

    repeating: bool
XPathType = XPath


class ODBCConnection:
    application: ApplicationType
    creator: XlCreator
    parent: object
    background_query: bool

    def cancel_refresh(self): pass

    command_text: object
    command_type: XlCmdType
    connection: object
    enable_refresh: bool

    def refresh(self): pass

    refresh_date: None
    refreshing: bool
    refresh_on_file_open: bool
    refresh_period: int
    robust_connect: XlRobustConnect

    def save_as_odc(self, odc_file_name: str, description: object = None, keywords: object = None): pass

    save_password: bool
    source_connection_file: str
    source_data: object
    source_data_file: str
    server_credentials_method: XlCredentialsMethod
    server_sso_application_id: str
    always_use_connection_file: bool
ODBCConnectionType = ODBCConnection


class AddIns:

    def __iter__(self, index: int) -> Iterable[AddInType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, filename: str, copy_file: object = None) -> AddInType: pass

    count: int

    def get(self, index: int) -> AddInType: pass

AddInsType = AddIns


class Outline:
    application: ApplicationType
    creator: XlCreator
    parent: object
    automatic_styles: bool

    def show_levels(self, row_levels: object = None, column_levels: object = None) -> object: pass

    summary_column: XlSummaryColumn
    summary_row: XlSummaryRow
OutlineType = Outline


class Tab:
    application: ApplicationType
    creator: XlCreator
    parent: object
    color: object
    color_index: XlColorIndex
    theme_color: XlThemeColor
    tint_and_shade: object
TabType = Tab


class SpellingOptions:
    dict_lang: int
    user_dict: str
    ignore_caps: bool
    suggest_main_only: bool
    ignore_mixed_digits: bool
    ignore_file_names: bool
    german_post_reform: bool
    korean_combine_aux: bool
    korean_use_auto_change_list: bool
    korean_process_compound: bool
    hebrew_modes: XlHebrewModes
    arabic_modes: XlArabicModes
    arabic_strict_alef_hamza: bool
    arabic_strict_final_yaa: bool
    arabic_strict_taa_marboota: bool
    russian_strict_e: bool
    spanish_modes: XlSpanishModes
    portugal_reform: XlPortugueseReform
    brazil_reform: XlPortugueseReform
SpellingOptionsType = SpellingOptions


class Actions:

    def __iter__(self, index: int) -> Iterable[ActionType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def get(self, index: int) -> ActionType: pass

ActionsType = Actions


class ProtectedViewWindows:

    def __iter__(self, index: int) -> Iterable[ProtectedViewWindowType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def open(self, filename: str, password: object = None, add_to_mru: object = None, repair_mode: object = None) -> ProtectedViewWindowType: pass


    def get(self, index: int) -> ProtectedViewWindowType: pass

ProtectedViewWindowsType = ProtectedViewWindows


class DropLines:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str

    def select(self) -> object: pass

    border: BorderType

    def delete(self) -> object: pass

    format: ChartFormatType
DropLinesType = DropLines


class XmlMap:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    is_exportable: bool
    show_import_export_validation_errors: bool
    save_data_source_definition: bool
    adjust_column_width: bool
    preserve_column_filter: bool
    preserve_number_formatting: bool
    append_on_import: bool
    root_element_name: str
    root_element_namespace: XmlNamespaceType
    schemas: XmlSchemasType
    data_binding: XmlDataBindingType

    def delete(self): pass


    def import_(self, url: str, overwrite: object = None) -> XlXmlImportResult: pass


    def import_xml(self, xml_data: str, overwrite: object = None) -> XlXmlImportResult: pass


    def export(self, url: str, overwrite: object = None) -> XlXmlExportResult: pass


    def export_xml(self, data: str) -> XlXmlExportResult: pass

    workbook_connection: WorkbookConnectionType
XmlMapType = XmlMap


class Interior:
    application: ApplicationType
    creator: XlCreator
    parent: object
    color: object
    color_index: object
    invert_if_negative: object
    pattern: object
    pattern_color: object
    pattern_color_index: object
    theme_color: object
    tint_and_shade: object
    pattern_theme_color: object
    pattern_tint_and_shade: object
    gradient: object
InteriorType = Interior


class PivotFilter:
    application: ApplicationType
    creator: XlCreator
    parent: object
    order: int
    filter_type: XlPivotFilterType
    name: str
    description: str

    def delete(self): pass

    active: bool
    pivot_field: PivotFieldType
    data_field: PivotFieldType
    data_cube_field: CubeFieldType
    value1: object
    value2: object
    member_property_field: PivotFieldType
    is_member_property_filter: bool
    whole_day_filter: bool
PivotFilterType = PivotFilter


class ChartColorFormat:
    application: ApplicationType
    creator: XlCreator
    parent: object
    scheme_color: int
    rgb: int
    type: int
ChartColorFormatType = ChartColorFormat


class Action:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str
    caption: str
    type: XlActionType
    coordinate: str
    content: str

    def execute(self): pass

ActionType = Action


class Parameters:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, name: str, i_data_type: object = None) -> ParameterType: pass

    count: int

    def item(self, index: object) -> ParameterType: pass


    def delete(self): pass

ParametersType = Parameters


class ModelFormatScientificNumber:
    application: ApplicationType
    creator: XlCreator
    parent: object
    decimal_places: int
ModelFormatScientificNumberType = ModelFormatScientificNumber


class ListColumns:

    def __iter__(self, index: int) -> Iterable[ListColumnType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object

    def add(self, position: object = None) -> ListColumnType: pass

    count: int

    def get(self, index: int) -> ListColumnType: pass

ListColumnsType = ListColumns


class PublishedDocs:
    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def item(self, index: int) -> PublishedDocType: pass

PublishedDocsType = PublishedDocs


class ModelFormatCurrency:
    application: ApplicationType
    creator: XlCreator
    parent: object
    symbol: str
    decimal_places: int
ModelFormatCurrencyType = ModelFormatCurrency


class Slicers:

    def __iter__(self, index: int) -> Iterable[SlicerType]: pass

    application: ApplicationType
    creator: XlCreator
    parent: object
    count: int

    def add(self, slicer_destination: object, level: object = None, name: object = None, caption: object = None, top: object = None, left: object = None, width: object = None, height: object = None) -> SlicerType: pass


    def get(self, index: int) -> SlicerType: pass

SlicersType = Slicers


class AutoCorrect:
    application: ApplicationType
    creator: XlCreator
    parent: object

    def add_replacement(self, what: str, replacement: str) -> object: pass

    capitalize_names_of_days: bool

    def delete_replacement(self, what: str) -> object: pass


    def get_replacement_list(self, index: object = None) -> object: pass


    def set_replacement_list(self, index: object = None, _param2: object = None): pass

    replace_text: bool
    two_initial_capitals: bool
    correct_sentence_cap: bool
    correct_caps_lock: bool
    display_auto_correct_options: bool
    auto_expand_list_range: bool
    auto_fill_formulas_in_lists: bool
AutoCorrectType = AutoCorrect


class Legend:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str

    def select(self) -> object: pass

    border: BorderType

    def delete(self) -> object: pass

    font: FontType

    def legend_entries(self, index: object = None) -> object: pass

    position: XlLegendPosition
    shadow: bool

    def clear(self) -> object: pass

    height: float
    interior: InteriorType
    fill: ChartFillFormatType
    left: float
    top: float
    width: float
    auto_scale_font: object
    include_in_layout: bool
    format: ChartFormatType

    def set_property(self, id: str, value: object): pass


    def get_property(self, id: str) -> object: pass

LegendType = Legend


class SeriesLines:
    application: ApplicationType
    creator: XlCreator
    parent: object
    name: str

    def select(self) -> object: pass

    border: BorderType

    def delete(self) -> object: pass

    format: ChartFormatType

    def set_property(self, id: str, value: object): pass


    def get_property(self, id: str) -> object: pass

SeriesLinesType = SeriesLines


