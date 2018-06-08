from enum import IntEnum

__all__ = ['XlActionType', 'XlAllocation', 'XlAllocationMethod', 'XlAllocationValue', 'XlApplyNamesOrder', 'XlArabicModes', 'XlArrangeStyle', 'XlAutoFillType', 'XlAutoFilterOperator', 'XlAxisGroup', 'XlBarShape', 'XlBinsType', 'XlBorderWeight', 'XlBordersIndex', 'XlBuiltInDialog', 'XlCalcMemNumberFormatType', 'XlCalculatedMemberType', 'XlCalculation', 'XlCalculationInterruptKey', 'XlCalculationState', 'XlCategoryLabelLevel', 'XlCellChangedState', 'XlCellInsertionMode', 'XlCellType', 'XlChartElementPosition', 'XlChartLocation', 'XlChartSplitType', 'XlChartType', 'XlCmdType', 'XlColorIndex', 'XlCommandUnderlines', 'XlCommentDisplayMode', 'XlConnectionType', 'XlConsolidationFunction', 'XlCopyPictureFormat', 'XlCreator', 'XlCredentialsMethod', 'XlCubeFieldSubType', 'XlCubeFieldType', 'XlCutCopyMode', 'XlDVType', 'XlDataLabelsType', 'XlDataSeriesDate', 'XlDataSeriesType', 'XlDirection', 'XlDisplayBlanksAs', 'XlDisplayDrawingObjects', 'XlEditionOptionsOption', 'XlEditionType', 'XlEnableCancelKey', 'XlEnableSelection', 'XlFileAccess', 'XlFileFormat', 'XlFileValidationPivotMode', 'XlFillWith', 'XlFilterAction', 'XlFilterStatus', 'XlFixedFormatType', 'XlFormControl', 'XlFormatConditionType', 'XlFormulaLabel', 'XlGenerateTableRefs', 'XlHAlign', 'XlHebrewModes', 'XlHtmlType', 'XlIconSet', 'XlLayoutFormType', 'XlLayoutRowType', 'XlLegendPosition', 'XlLinkInfo', 'XlLinkType', 'XlListConflict', 'XlListDataType', 'XlListObjectSourceType', 'XlLocationInTable', 'XlMSApplication', 'XlMailSystem', 'XlMousePointer', 'XlOartHorizontalOverflow', 'XlOartVerticalOverflow', 'XlObjectSize', 'XlOrder', 'XlOrientation', 'XlPTSelectionMode', 'XlPageBreak', 'XlPageBreakExtent', 'XlPageOrientation', 'XlPaperSize', 'XlParameterDataType', 'XlParameterType', 'XlPasteSpecialOperation', 'XlPasteType', 'XlPictureAppearance', 'XlPivotCellType', 'XlPivotFieldCalculation', 'XlPivotFieldDataType', 'XlPivotFieldOrientation', 'XlPivotFieldRepeatLabels', 'XlPivotFilterType', 'XlPivotFormatType', 'XlPivotLineType', 'XlPivotTableMissingItems', 'XlPivotTableSourceType', 'XlPivotTableVersionList', 'XlPlacement', 'XlPlatform', 'XlPortugueseReform', 'XlPrintErrors', 'XlPrintLocation', 'XlPriority', 'XlProtectedViewWindowState', 'XlPublishToDocsDisclosureScope', 'XlQueryType', 'XlQuickAnalysisMode', 'XlRangeAutoFormat', 'XlReferenceStyle', 'XlRemoveDocInfoType', 'XlRobustConnect', 'XlRoutingSlipDelivery', 'XlRoutingSlipStatus', 'XlRowCol', 'XlRunAutoMacro', 'XlSaveAsAccessMode', 'XlSaveConflictResolution', 'XlSearchDirection', 'XlSeriesNameLevel', 'XlSheetType', 'XlSheetVisibility', 'XlSizeRepresents', 'XlSlicerCacheType', 'XlSlicerCrossFilterType', 'XlSlicerSort', 'XlSmartTagControlType', 'XlSmartTagDisplayMode', 'XlSortDataOption', 'XlSortMethod', 'XlSortOn', 'XlSortOrder', 'XlSortOrientation', 'XlSourceType', 'XlSpanishModes', 'XlSparkScale', 'XlSparkType', 'XlSparklineRowCol', 'XlSpeakDirection', 'XlSubscribeToFormat', 'XlSubtototalLocationType', 'XlSummaryColumn', 'XlSummaryRow', 'XlTableStyleElementType', 'XlTextParsingType', 'XlTextQualifier', 'XlTextVisualLayoutType', 'XlThemeColor', 'XlThemeFont', 'XlThreadMode', 'XlTickLabelOrientation', 'XlTimelineLevel', 'XlToolbarProtection', 'XlTotalsCalculation', 'XlUpdateLinks', 'XlVAlign', 'XlWebFormatting', 'XlWebSelectionType', 'XlWindowState', 'XlWindowType', 'XlWindowView', 'XlXLMMacroType', 'XlXmlExportResult', 'XlXmlImportResult', 'XlYesNoGuess']


class XlBorderWeight(IntEnum):
    xlMedium = -4138
    xlHairline = 1
    xlThin = 2
    xlThick = 4


class XlSlicerSort(IntEnum):
    xlSlicerSortDataSourceOrder = 1
    xlSlicerSortAscending = 2
    xlSlicerSortDescending = 3


class XlSpeakDirection(IntEnum):
    xlSpeakByRows = 0
    xlSpeakByColumns = 1


class XlFileFormat(IntEnum):
    xlCurrentPlatformText = -4158
    xlWorkbookNormal = -4143
    xlSYLK = 2
    xlWKS = 4
    xlWK1 = 5
    xlCSV = 6
    xlDBF2 = 7
    xlDBF3 = 8
    xlDIF = 9
    xlDBF4 = 11
    xlWJ2WD1 = 14
    xlWK3 = 15
    xlExcel2 = 16
    xlTemplate = 17
    xlTemplate8 = 17
    xlAddIn = 18
    xlAddIn8 = 18
    xlTextMac = 19
    xlTextWindows = 20
    xlTextMSDOS = 21
    xlCSVMac = 22
    xlCSVWindows = 23
    xlCSVMSDOS = 24
    xlIntlMacro = 25
    xlIntlAddIn = 26
    xlExcel2FarEast = 27
    xlWorks2FarEast = 28
    xlExcel3 = 29
    xlWK1FMT = 30
    xlWK1ALL = 31
    xlWK3FM3 = 32
    xlExcel4 = 33
    xlWQ1 = 34
    xlExcel4Workbook = 35
    xlTextPrinter = 36
    xlWK4 = 38
    xlExcel5 = 39
    xlExcel7 = 39
    xlWJ3 = 40
    xlWJ3FJ3 = 41
    xlUnicodeText = 42
    xlExcel9795 = 43
    xlHtml = 44
    xlWebArchive = 45
    xlXMLSpreadsheet = 46
    xlExcel12 = 50
    xlOpenXMLWorkbook = 51
    xlWorkbookDefault = 51
    xlOpenXMLWorkbookMacroEnabled = 52
    xlOpenXMLTemplateMacroEnabled = 53
    xlOpenXMLTemplate = 54
    xlOpenXMLAddIn = 55
    xlExcel8 = 56
    xlOpenDocumentSpreadsheet = 60
    xlOpenXMLStrictWorkbook = 61
    xlCSVUTF8 = 62


class XlSearchDirection(IntEnum):
    xlNext = 1
    xlPrevious = 2


class XlPasteType(IntEnum):
    xlPasteValues = -4163
    xlPasteComments = -4144
    xlPasteFormulas = -4123
    xlPasteFormats = -4122
    xlPasteAll = -4104
    xlPasteValidation = 6
    xlPasteAllExceptBorders = 7
    xlPasteColumnWidths = 8
    xlPasteFormulasAndNumberFormats = 11
    xlPasteValuesAndNumberFormats = 12
    xlPasteAllUsingSourceTheme = 13
    xlPasteAllMergingConditionalFormats = 14


class XlPivotFieldCalculation(IntEnum):
    xlNoAdditionalCalculation = -4143
    xlDifferenceFrom = 2
    xlPercentOf = 3
    xlPercentDifferenceFrom = 4
    xlRunningTotal = 5
    xlPercentOfRow = 6
    xlPercentOfColumn = 7
    xlPercentOfTotal = 8
    xlIndex = 9
    xlPercentOfParentRow = 10
    xlPercentOfParentColumn = 11
    xlPercentOfParent = 12
    xlPercentRunningTotal = 13
    xlRankAscending = 14
    xlRankDecending = 15


class XlSortOrientation(IntEnum):
    xlSortColumns = 1
    xlSortRows = 2


class XlChartSplitType(IntEnum):
    xlSplitByPosition = 1
    xlSplitByValue = 2
    xlSplitByPercentValue = 3
    xlSplitByCustomSplit = 4


class XlOrientation(IntEnum):
    xlUpward = -4171
    xlDownward = -4170
    xlVertical = -4166
    xlHorizontal = -4128


class XlQuickAnalysisMode(IntEnum):
    xlLensOnly = 0
    xlFormatConditions = 1
    xlRecommendedCharts = 2
    xlTotals = 3
    xlTables = 4
    xlSparklines = 5


class XlSmartTagControlType(IntEnum):
    xlSmartTagControlSmartTag = 1
    xlSmartTagControlLink = 2
    xlSmartTagControlHelp = 3
    xlSmartTagControlHelpURL = 4
    xlSmartTagControlSeparator = 5
    xlSmartTagControlButton = 6
    xlSmartTagControlLabel = 7
    xlSmartTagControlImage = 8
    xlSmartTagControlCheckbox = 9
    xlSmartTagControlTextbox = 10
    xlSmartTagControlListbox = 11
    xlSmartTagControlCombo = 12
    xlSmartTagControlActiveX = 13
    xlSmartTagControlRadioGroup = 14


class XlParameterType(IntEnum):
    xlPrompt = 0
    xlConstant = 1
    xlRange = 2


class XlPaperSize(IntEnum):
    xlPaperLetter = 1
    xlPaperLetterSmall = 2
    xlPaperTabloid = 3
    xlPaperLedger = 4
    xlPaperLegal = 5
    xlPaperStatement = 6
    xlPaperExecutive = 7
    xlPaperA3 = 8
    xlPaperA4 = 9
    xlPaperA4Small = 10
    xlPaperA5 = 11
    xlPaperB4 = 12
    xlPaperB5 = 13
    xlPaperFolio = 14
    xlPaperQuarto = 15
    xlPaper10x14 = 16
    xlPaper11x17 = 17
    xlPaperNote = 18
    xlPaperEnvelope9 = 19
    xlPaperEnvelope10 = 20
    xlPaperEnvelope11 = 21
    xlPaperEnvelope12 = 22
    xlPaperEnvelope14 = 23
    xlPaperCsheet = 24
    xlPaperDsheet = 25
    xlPaperEsheet = 26
    xlPaperEnvelopeDL = 27
    xlPaperEnvelopeC5 = 28
    xlPaperEnvelopeC3 = 29
    xlPaperEnvelopeC4 = 30
    xlPaperEnvelopeC6 = 31
    xlPaperEnvelopeC65 = 32
    xlPaperEnvelopeB4 = 33
    xlPaperEnvelopeB5 = 34
    xlPaperEnvelopeB6 = 35
    xlPaperEnvelopeItaly = 36
    xlPaperEnvelopeMonarch = 37
    xlPaperEnvelopePersonal = 38
    xlPaperFanfoldUS = 39
    xlPaperFanfoldStdGerman = 40
    xlPaperFanfoldLegalGerman = 41
    xlPaperUser = 256


class XlBordersIndex(IntEnum):
    xlDiagonalDown = 5
    xlDiagonalUp = 6
    xlEdgeLeft = 7
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlEdgeRight = 10
    xlInsideVertical = 11
    xlInsideHorizontal = 12


class XlWindowType(IntEnum):
    xlInfo = -4129
    xlWorkbook = 1
    xlClipboard = 3
    xlChartInPlace = 4
    xlChartAsWindow = 5


class XlActionType(IntEnum):
    xlActionTypeUrl = 1
    xlActionTypeRowset = 16
    xlActionTypeReport = 128
    xlActionTypeDrillthrough = 256


class XlPrintErrors(IntEnum):
    xlPrintErrorsDisplayed = 0
    xlPrintErrorsBlank = 1
    xlPrintErrorsDash = 2
    xlPrintErrorsNA = 3


class XlHtmlType(IntEnum):
    xlHtmlStatic = 0
    xlHtmlCalc = 1
    xlHtmlList = 2
    xlHtmlChart = 3


class XlConnectionType(IntEnum):
    xlConnectionTypeOLEDB = 1
    xlConnectionTypeODBC = 2
    xlConnectionTypeXMLMAP = 3
    xlConnectionTypeTEXT = 4
    xlConnectionTypeWEB = 5
    xlConnectionTypeDATAFEED = 6
    xlConnectionTypeMODEL = 7
    xlConnectionTypeWORKSHEET = 8
    xlConnectionTypeNOSOURCE = 9


class XlCredentialsMethod(IntEnum):
    xlCredentialsMethodIntegrated = 0
    xlCredentialsMethodNone = 1
    xlCredentialsMethodStored = 2


class XlDVType(IntEnum):
    xlValidateInputOnly = 0
    xlValidateWholeNumber = 1
    xlValidateDecimal = 2
    xlValidateList = 3
    xlValidateDate = 4
    xlValidateTime = 5
    xlValidateTextLength = 6
    xlValidateCustom = 7


class XlPageBreak(IntEnum):
    xlPageBreakNone = -4142
    xlPageBreakManual = -4135
    xlPageBreakAutomatic = -4105


class XlToolbarProtection(IntEnum):
    xlToolbarProtectionNone = -4143
    xlNoButtonChanges = 1
    xlNoShapeChanges = 2
    xlNoDockingChanges = 3
    xlNoChanges = 4


class XlTextVisualLayoutType(IntEnum):
    xlTextVisualLTR = 1
    xlTextVisualRTL = 2


class XlSlicerCrossFilterType(IntEnum):
    xlSlicerNoCrossFilter = 1
    xlSlicerCrossFilterShowItemsWithDataAtTop = 2
    xlSlicerCrossFilterShowItemsWithNoData = 3
    xlSlicerCrossFilterHideButtonsWithNoData = 4


class XlOartHorizontalOverflow(IntEnum):
    xlOartHorizontalOverflowOverflow = 0
    xlOartHorizontalOverflowClip = 1


class XlDisplayBlanksAs(IntEnum):
    xlNotPlotted = 1
    xlZero = 2
    xlInterpolated = 3


class XlPriority(IntEnum):
    xlPriorityNormal = -4143
    xlPriorityLow = -4134
    xlPriorityHigh = -4127


class XlSheetVisibility(IntEnum):
    xlSheetVisible = -1
    xlSheetHidden = 0
    xlSheetVeryHidden = 2


class XlWindowView(IntEnum):
    xlNormalView = 1
    xlPageBreakPreview = 2
    xlPageLayoutView = 3


class XlCategoryLabelLevel(IntEnum):
    xlCategoryLabelLevelNone = -3
    xlCategoryLabelLevelCustom = -2
    xlCategoryLabelLevelAll = -1


class XlSaveAsAccessMode(IntEnum):
    xlNoChange = 1
    xlShared = 2
    xlExclusive = 3


class XlWebSelectionType(IntEnum):
    xlEntirePage = 1
    xlAllTables = 2
    xlSpecifiedTables = 3


class XlOartVerticalOverflow(IntEnum):
    xlOartVerticalOverflowOverflow = 0
    xlOartVerticalOverflowClip = 1
    xlOartVerticalOverflowEllipsis = 2


class XlLayoutFormType(IntEnum):
    xlTabular = 0
    xlOutline = 1


class XlTextQualifier(IntEnum):
    xlTextQualifierNone = -4142
    xlTextQualifierDoubleQuote = 1
    xlTextQualifierSingleQuote = 2


class XlFilterStatus(IntEnum):
    xlFilterStatusOK = 0
    xlFilterStatusDateWrongOrder = 1
    xlFilterStatusDateHasTime = 2
    xlFilterStatusInvalidDate = 3


class XlTotalsCalculation(IntEnum):
    xlTotalsCalculationNone = 0
    xlTotalsCalculationSum = 1
    xlTotalsCalculationAverage = 2
    xlTotalsCalculationCount = 3
    xlTotalsCalculationCountNums = 4
    xlTotalsCalculationMin = 5
    xlTotalsCalculationMax = 6
    xlTotalsCalculationStdDev = 7
    xlTotalsCalculationVar = 8
    xlTotalsCalculationCustom = 9


class XlBarShape(IntEnum):
    xlBox = 0
    xlPyramidToPoint = 1
    xlPyramidToMax = 2
    xlCylinder = 3
    xlConeToPoint = 4
    xlConeToMax = 5


class XlChartType(IntEnum):
    xlXYScatter = -4169
    xlRadar = -4151
    xlDoughnut = -4120
    xl3DPie = -4102
    xl3DLine = -4101
    xl3DColumn = -4100
    xl3DArea = -4098
    xlArea = 1
    xlLine = 4
    xlPie = 5
    xlBubble = 15
    xlColumnClustered = 51
    xlColumnStacked = 52
    xlColumnStacked100 = 53
    xl3DColumnClustered = 54
    xl3DColumnStacked = 55
    xl3DColumnStacked100 = 56
    xlBarClustered = 57
    xlBarStacked = 58
    xlBarStacked100 = 59
    xl3DBarClustered = 60
    xl3DBarStacked = 61
    xl3DBarStacked100 = 62
    xlLineStacked = 63
    xlLineStacked100 = 64
    xlLineMarkers = 65
    xlLineMarkersStacked = 66
    xlLineMarkersStacked100 = 67
    xlPieOfPie = 68
    xlPieExploded = 69
    xl3DPieExploded = 70
    xlBarOfPie = 71
    xlXYScatterSmooth = 72
    xlXYScatterSmoothNoMarkers = 73
    xlXYScatterLines = 74
    xlXYScatterLinesNoMarkers = 75
    xlAreaStacked = 76
    xlAreaStacked100 = 77
    xl3DAreaStacked = 78
    xl3DAreaStacked100 = 79
    xlDoughnutExploded = 80
    xlRadarMarkers = 81
    xlRadarFilled = 82
    xlSurface = 83
    xlSurfaceWireframe = 84
    xlSurfaceTopView = 85
    xlSurfaceTopViewWireframe = 86
    xlBubble3DEffect = 87
    xlStockHLC = 88
    xlStockOHLC = 89
    xlStockVHLC = 90
    xlStockVOHLC = 91
    xlCylinderColClustered = 92
    xlCylinderColStacked = 93
    xlCylinderColStacked100 = 94
    xlCylinderBarClustered = 95
    xlCylinderBarStacked = 96
    xlCylinderBarStacked100 = 97
    xlCylinderCol = 98
    xlConeColClustered = 99
    xlConeColStacked = 100
    xlConeColStacked100 = 101
    xlConeBarClustered = 102
    xlConeBarStacked = 103
    xlConeBarStacked100 = 104
    xlConeCol = 105
    xlPyramidColClustered = 106
    xlPyramidColStacked = 107
    xlPyramidColStacked100 = 108
    xlPyramidBarClustered = 109
    xlPyramidBarStacked = 110
    xlPyramidBarStacked100 = 111
    xlPyramidCol = 112
    xlTreemap = 117
    xlHistogram = 118
    xlWaterfall = 119
    xlSunburst = 120
    xlBoxwhisker = 121
    xlPareto = 122
    xlFunnel = 123
    xlRegionMap = 140


class XlThreadMode(IntEnum):
    xlThreadModeAutomatic = 0
    xlThreadModeManual = 1


class XlProtectedViewWindowState(IntEnum):
    xlProtectedViewWindowNormal = 0
    xlProtectedViewWindowMinimized = 1
    xlProtectedViewWindowMaximized = 2


class XlTextParsingType(IntEnum):
    xlDelimited = 1
    xlFixedWidth = 2


class XlCubeFieldType(IntEnum):
    xlHierarchy = 1
    xlMeasure = 2
    xlSet = 3


class XlCreator(IntEnum):
    xlCreatorCode = 1480803660


class XlFillWith(IntEnum):
    xlFillWithFormats = -4122
    xlFillWithAll = -4104
    xlFillWithContents = 2


class XlApplyNamesOrder(IntEnum):
    xlRowThenColumn = 1
    xlColumnThenRow = 2


class XlLocationInTable(IntEnum):
    xlRowHeader = -4153
    xlColumnHeader = -4110
    xlPageHeader = 2
    xlDataHeader = 3
    xlRowItem = 4
    xlColumnItem = 5
    xlPageItem = 6
    xlDataItem = 7
    xlTableBody = 8


class XlPTSelectionMode(IntEnum):
    xlDataAndLabel = 0
    xlLabelOnly = 1
    xlDataOnly = 2
    xlOrigin = 3
    xlBlanks = 4
    xlButton = 15
    xlFirstRow = 256


class XlSummaryColumn(IntEnum):
    xlSummaryOnRight = -4152
    xlSummaryOnLeft = -4131


class XlCubeFieldSubType(IntEnum):
    xlCubeHierarchy = 1
    xlCubeMeasure = 2
    xlCubeSet = 3
    xlCubeAttribute = 4
    xlCubeCalculatedMeasure = 5
    xlCubeKPIValue = 6
    xlCubeKPIGoal = 7
    xlCubeKPIStatus = 8
    xlCubeKPITrend = 9
    xlCubeKPIWeight = 10
    xlCubeImplicitMeasure = 11


class XlCommentDisplayMode(IntEnum):
    xlCommentIndicatorOnly = -1
    xlNoIndicator = 0
    xlCommentAndIndicator = 1


class XlWebFormatting(IntEnum):
    xlWebFormattingAll = 1
    xlWebFormattingRTF = 2
    xlWebFormattingNone = 3


class XlDisplayDrawingObjects(IntEnum):
    xlDisplayShapes = -4104
    xlPlaceholders = 2
    xlHide = 3


class XlReferenceStyle(IntEnum):
    xlR1C1 = -4150
    xlA1 = 1


class XlSmartTagDisplayMode(IntEnum):
    xlIndicatorAndButton = 0
    xlDisplayNone = 1
    xlButtonOnly = 2


class XlAllocationValue(IntEnum):
    xlAllocateValue = 1
    xlAllocateIncrement = 2


class XlMousePointer(IntEnum):
    xlDefault = -4143
    xlNorthwestArrow = 1
    xlWait = 2
    xlIBeam = 3


class XlLayoutRowType(IntEnum):
    xlCompactRow = 0
    xlTabularRow = 1
    xlOutlineRow = 2


class XlSortDataOption(IntEnum):
    xlSortNormal = 0
    xlSortTextAsNumbers = 1


class XlFilterAction(IntEnum):
    xlFilterInPlace = 1
    xlFilterCopy = 2


class XlBinsType(IntEnum):
    xlBinsTypeAutomatic = 0
    xlBinsTypeCategorical = 1
    xlBinsTypeManual = 2
    xlBinsTypeBinSize = 3
    xlBinsTypeBinCount = 4


class XlAllocationMethod(IntEnum):
    xlEqualAllocation = 1
    xlWeightedAllocation = 2


class XlPivotFieldDataType(IntEnum):
    xlText = -4158
    xlNumber = -4145
    xlDate = 2


class XlMailSystem(IntEnum):
    xlNoMailSystem = 0
    xlMAPI = 1
    xlPowerTalk = 2


class XlHAlign(IntEnum):
    xlHAlignRight = -4152
    xlHAlignLeft = -4131
    xlHAlignJustify = -4130
    xlHAlignDistributed = -4117
    xlHAlignCenter = -4108
    xlHAlignGeneral = 1
    xlHAlignFill = 5
    xlHAlignCenterAcrossSelection = 7


class XlAxisGroup(IntEnum):
    xlPrimary = 1
    xlSecondary = 2


class XlRobustConnect(IntEnum):
    xlAsRequired = 0
    xlAlways = 1
    xlNever = 2


class XlPivotTableMissingItems(IntEnum):
    xlMissingItemsDefault = -1
    xlMissingItemsNone = 0
    xlMissingItemsMax = 32500
    xlMissingItemsMax2 = 1048576


class XlThemeColor(IntEnum):
    xlThemeColorDark1 = 1
    xlThemeColorLight1 = 2
    xlThemeColorDark2 = 3
    xlThemeColorLight2 = 4
    xlThemeColorAccent1 = 5
    xlThemeColorAccent2 = 6
    xlThemeColorAccent3 = 7
    xlThemeColorAccent4 = 8
    xlThemeColorAccent5 = 9
    xlThemeColorAccent6 = 10
    xlThemeColorHyperlink = 11
    xlThemeColorFollowedHyperlink = 12


class XlParameterDataType(IntEnum):
    xlParamTypeWChar = -8
    xlParamTypeBit = -7
    xlParamTypeTinyInt = -6
    xlParamTypeBigInt = -5
    xlParamTypeLongVarBinary = -4
    xlParamTypeVarBinary = -3
    xlParamTypeBinary = -2
    xlParamTypeLongVarChar = -1
    xlParamTypeUnknown = 0
    xlParamTypeChar = 1
    xlParamTypeNumeric = 2
    xlParamTypeDecimal = 3
    xlParamTypeInteger = 4
    xlParamTypeSmallInt = 5
    xlParamTypeFloat = 6
    xlParamTypeReal = 7
    xlParamTypeDouble = 8
    xlParamTypeDate = 9
    xlParamTypeTime = 10
    xlParamTypeTimestamp = 11
    xlParamTypeVarChar = 12


class XlConsolidationFunction(IntEnum):
    xlVarP = -4165
    xlVar = -4164
    xlSum = -4157
    xlStDevP = -4156
    xlStDev = -4155
    xlProduct = -4149
    xlMin = -4139
    xlMax = -4136
    xlCountNums = -4113
    xlCount = -4112
    xlAverage = -4106
    xlDistinctCount = 11
    xlUnknown = 1000


class XlDataLabelsType(IntEnum):
    xlDataLabelsShowNone = -4142
    xlDataLabelsShowValue = 2
    xlDataLabelsShowPercent = 3
    xlDataLabelsShowLabel = 4
    xlDataLabelsShowLabelAndPercent = 5
    xlDataLabelsShowBubbleSizes = 6


class XlOrder(IntEnum):
    xlDownThenOver = 1
    xlOverThenDown = 2


class XlSparkType(IntEnum):
    xlSparkLine = 1
    xlSparkColumn = 2
    xlSparkColumnStacked100 = 3


class XlObjectSize(IntEnum):
    xlScreenSize = 1
    xlFitToPage = 2
    xlFullPage = 3


class XlRunAutoMacro(IntEnum):
    xlAutoOpen = 1
    xlAutoClose = 2
    xlAutoActivate = 3
    xlAutoDeactivate = 4


class XlArrangeStyle(IntEnum):
    xlArrangeStyleVertical = -4166
    xlArrangeStyleHorizontal = -4128
    xlArrangeStyleTiled = 1
    xlArrangeStyleCascade = 7


class XlPasteSpecialOperation(IntEnum):
    xlPasteSpecialOperationNone = -4142
    xlPasteSpecialOperationAdd = 2
    xlPasteSpecialOperationSubtract = 3
    xlPasteSpecialOperationMultiply = 4
    xlPasteSpecialOperationDivide = 5


class XlPivotFilterType(IntEnum):
    xlTopCount = 1
    xlBottomCount = 2
    xlTopPercent = 3
    xlBottomPercent = 4
    xlTopSum = 5
    xlBottomSum = 6
    xlValueEquals = 7
    xlValueDoesNotEqual = 8
    xlValueIsGreaterThan = 9
    xlValueIsGreaterThanOrEqualTo = 10
    xlValueIsLessThan = 11
    xlValueIsLessThanOrEqualTo = 12
    xlValueIsBetween = 13
    xlValueIsNotBetween = 14
    xlCaptionEquals = 15
    xlCaptionDoesNotEqual = 16
    xlCaptionBeginsWith = 17
    xlCaptionDoesNotBeginWith = 18
    xlCaptionEndsWith = 19
    xlCaptionDoesNotEndWith = 20
    xlCaptionContains = 21
    xlCaptionDoesNotContain = 22
    xlCaptionIsGreaterThan = 23
    xlCaptionIsGreaterThanOrEqualTo = 24
    xlCaptionIsLessThan = 25
    xlCaptionIsLessThanOrEqualTo = 26
    xlCaptionIsBetween = 27
    xlCaptionIsNotBetween = 28
    xlSpecificDate = 29
    xlNotSpecificDate = 30
    xlBefore = 31
    xlBeforeOrEqualTo = 32
    xlAfter = 33
    xlAfterOrEqualTo = 34
    xlDateBetween = 35
    xlDateNotBetween = 36
    xlDateTomorrow = 37
    xlDateToday = 38
    xlDateYesterday = 39
    xlDateNextWeek = 40
    xlDateThisWeek = 41
    xlDateLastWeek = 42
    xlDateNextMonth = 43
    xlDateThisMonth = 44
    xlDateLastMonth = 45
    xlDateNextQuarter = 46
    xlDateThisQuarter = 47
    xlDateLastQuarter = 48
    xlDateNextYear = 49
    xlDateThisYear = 50
    xlDateLastYear = 51
    xlYearToDate = 52
    xlAllDatesInPeriodQuarter1 = 53
    xlAllDatesInPeriodQuarter2 = 54
    xlAllDatesInPeriodQuarter3 = 55
    xlAllDatesInPeriodQuarter4 = 56
    xlAllDatesInPeriodJanuary = 57
    xlAllDatesInPeriodFebruary = 58
    xlAllDatesInPeriodMarch = 59
    xlAllDatesInPeriodApril = 60
    xlAllDatesInPeriodMay = 61
    xlAllDatesInPeriodJune = 62
    xlAllDatesInPeriodJuly = 63
    xlAllDatesInPeriodAugust = 64
    xlAllDatesInPeriodSeptember = 65
    xlAllDatesInPeriodOctober = 66
    xlAllDatesInPeriodNovember = 67
    xlAllDatesInPeriodDecember = 68


class XlPageOrientation(IntEnum):
    xlPortrait = 1
    xlLandscape = 2


class XlColorIndex(IntEnum):
    xlColorIndexNone = -4142
    xlColorIndexAutomatic = -4105


class XlCellChangedState(IntEnum):
    xlCellNotChanged = 1
    xlCellChanged = 2
    xlCellChangeApplied = 3


class XlCalculationState(IntEnum):
    xlDone = 0
    xlCalculating = 1
    xlPending = 2


class XlSheetType(IntEnum):
    xlWorksheet = -4167
    xlDialogSheet = -4116
    xlChart = -4109
    xlExcel4MacroSheet = 3
    xlExcel4IntlMacroSheet = 4


class XlRowCol(IntEnum):
    xlRows = 1
    xlColumns = 2


class XlSourceType(IntEnum):
    xlSourceWorkbook = 0
    xlSourceSheet = 1
    xlSourcePrintArea = 2
    xlSourceAutoFilter = 3
    xlSourceRange = 4
    xlSourceChart = 5
    xlSourcePivotTable = 6
    xlSourceQuery = 7


class XlEditionOptionsOption(IntEnum):
    xlCancel = 1
    xlSendPublisher = 2
    xlUpdateSubscriber = 2
    xlOpenSource = 3
    xlSelect = 3
    xlAutomaticUpdate = 4
    xlManualUpdate = 5
    xlChangeAttributes = 6


class XlBuiltInDialog(IntEnum):
    xlDialogOpen = 1
    xlDialogOpenLinks = 2
    xlDialogSaveAs = 5
    xlDialogFileDelete = 6
    xlDialogPageSetup = 7
    xlDialogPrint = 8
    xlDialogPrinterSetup = 9
    xlDialogArrangeAll = 12
    xlDialogWindowSize = 13
    xlDialogWindowMove = 14
    xlDialogRun = 17
    xlDialogSetPrintTitles = 23
    xlDialogFont = 26
    xlDialogDisplay = 27
    xlDialogProtectDocument = 28
    xlDialogCalculation = 32
    xlDialogExtract = 35
    xlDialogDataDelete = 36
    xlDialogSort = 39
    xlDialogDataSeries = 40
    xlDialogTable = 41
    xlDialogFormatNumber = 42
    xlDialogAlignment = 43
    xlDialogStyle = 44
    xlDialogBorder = 45
    xlDialogCellProtection = 46
    xlDialogColumnWidth = 47
    xlDialogClear = 52
    xlDialogPasteSpecial = 53
    xlDialogEditDelete = 54
    xlDialogInsert = 55
    xlDialogPasteNames = 58
    xlDialogDefineName = 61
    xlDialogCreateNames = 62
    xlDialogFormulaGoto = 63
    xlDialogFormulaFind = 64
    xlDialogGalleryArea = 67
    xlDialogGalleryBar = 68
    xlDialogGalleryColumn = 69
    xlDialogGalleryLine = 70
    xlDialogGalleryPie = 71
    xlDialogGalleryScatter = 72
    xlDialogCombination = 73
    xlDialogGridlines = 76
    xlDialogAxes = 78
    xlDialogAttachText = 80
    xlDialogPatterns = 84
    xlDialogMainChart = 85
    xlDialogOverlay = 86
    xlDialogScale = 87
    xlDialogFormatLegend = 88
    xlDialogFormatText = 89
    xlDialogParse = 91
    xlDialogUnhide = 94
    xlDialogWorkspace = 95
    xlDialogActivate = 103
    xlDialogCopyPicture = 108
    xlDialogDeleteName = 110
    xlDialogDeleteFormat = 111
    xlDialogNew = 119
    xlDialogRowHeight = 127
    xlDialogFormatMove = 128
    xlDialogFormatSize = 129
    xlDialogFormulaReplace = 130
    xlDialogSelectSpecial = 132
    xlDialogApplyNames = 133
    xlDialogReplaceFont = 134
    xlDialogSplit = 137
    xlDialogOutline = 142
    xlDialogSaveWorkbook = 145
    xlDialogCopyChart = 147
    xlDialogFormatFont = 150
    xlDialogNote = 154
    xlDialogSetUpdateStatus = 159
    xlDialogColorPalette = 161
    xlDialogChangeLink = 166
    xlDialogAppMove = 170
    xlDialogAppSize = 171
    xlDialogMainChartType = 185
    xlDialogOverlayChartType = 186
    xlDialogOpenMail = 188
    xlDialogSendMail = 189
    xlDialogStandardFont = 190
    xlDialogConsolidate = 191
    xlDialogSortSpecial = 192
    xlDialogGallery3dArea = 193
    xlDialogGallery3dColumn = 194
    xlDialogGallery3dLine = 195
    xlDialogGallery3dPie = 196
    xlDialogView3d = 197
    xlDialogGoalSeek = 198
    xlDialogWorkgroup = 199
    xlDialogFillGroup = 200
    xlDialogUpdateLink = 201
    xlDialogPromote = 202
    xlDialogDemote = 203
    xlDialogShowDetail = 204
    xlDialogObjectProperties = 207
    xlDialogSaveNewObject = 208
    xlDialogApplyStyle = 212
    xlDialogAssignToObject = 213
    xlDialogObjectProtection = 214
    xlDialogCreatePublisher = 217
    xlDialogSubscribeTo = 218
    xlDialogShowToolbar = 220
    xlDialogPrintPreview = 222
    xlDialogEditColor = 223
    xlDialogFormatMain = 225
    xlDialogFormatOverlay = 226
    xlDialogEditSeries = 228
    xlDialogDefineStyle = 229
    xlDialogGalleryRadar = 249
    xlDialogEditionOptions = 251
    xlDialogZoom = 256
    xlDialogInsertObject = 259
    xlDialogSize = 261
    xlDialogMove = 262
    xlDialogFormatAuto = 269
    xlDialogGallery3dBar = 272
    xlDialogGallery3dSurface = 273
    xlDialogCustomizeToolbar = 276
    xlDialogWorkbookAdd = 281
    xlDialogWorkbookMove = 282
    xlDialogWorkbookCopy = 283
    xlDialogWorkbookOptions = 284
    xlDialogSaveWorkspace = 285
    xlDialogChartWizard = 288
    xlDialogAssignToTool = 293
    xlDialogPlacement = 300
    xlDialogFillWorkgroup = 301
    xlDialogWorkbookNew = 302
    xlDialogScenarioCells = 305
    xlDialogScenarioAdd = 307
    xlDialogScenarioEdit = 308
    xlDialogScenarioSummary = 311
    xlDialogPivotTableWizard = 312
    xlDialogPivotFieldProperties = 313
    xlDialogOptionsCalculation = 318
    xlDialogOptionsEdit = 319
    xlDialogOptionsView = 320
    xlDialogAddinManager = 321
    xlDialogMenuEditor = 322
    xlDialogAttachToolbars = 323
    xlDialogOptionsChart = 325
    xlDialogVbaInsertFile = 328
    xlDialogVbaProcedureDefinition = 330
    xlDialogRoutingSlip = 336
    xlDialogMailLogon = 339
    xlDialogInsertPicture = 342
    xlDialogGalleryDoughnut = 344
    xlDialogChartTrend = 350
    xlDialogWorkbookInsert = 354
    xlDialogOptionsTransition = 355
    xlDialogOptionsGeneral = 356
    xlDialogFilterAdvanced = 370
    xlDialogMailNextLetter = 378
    xlDialogDataLabel = 379
    xlDialogInsertTitle = 380
    xlDialogFontProperties = 381
    xlDialogMacroOptions = 382
    xlDialogWorkbookUnhide = 384
    xlDialogWorkbookName = 386
    xlDialogGalleryCustom = 388
    xlDialogAddChartAutoformat = 390
    xlDialogChartAddData = 392
    xlDialogTabOrder = 394
    xlDialogSubtotalCreate = 398
    xlDialogWorkbookTabSplit = 415
    xlDialogWorkbookProtect = 417
    xlDialogScrollbarProperties = 420
    xlDialogPivotShowPages = 421
    xlDialogTextToColumns = 422
    xlDialogFormatCharttype = 423
    xlDialogPivotFieldGroup = 433
    xlDialogPivotFieldUngroup = 434
    xlDialogCheckboxProperties = 435
    xlDialogLabelProperties = 436
    xlDialogListboxProperties = 437
    xlDialogEditboxProperties = 438
    xlDialogOpenText = 441
    xlDialogPushbuttonProperties = 445
    xlDialogFilter = 447
    xlDialogFunctionWizard = 450
    xlDialogSaveCopyAs = 456
    xlDialogOptionsListsAdd = 458
    xlDialogSeriesAxes = 460
    xlDialogSeriesX = 461
    xlDialogSeriesY = 462
    xlDialogErrorbarX = 463
    xlDialogErrorbarY = 464
    xlDialogFormatChart = 465
    xlDialogSeriesOrder = 466
    xlDialogMailEditMailer = 470
    xlDialogStandardWidth = 472
    xlDialogScenarioMerge = 473
    xlDialogProperties = 474
    xlDialogSummaryInfo = 474
    xlDialogFindFile = 475
    xlDialogActiveCellFont = 476
    xlDialogVbaMakeAddin = 478
    xlDialogFileSharing = 481
    xlDialogAutoCorrect = 485
    xlDialogCustomViews = 493
    xlDialogInsertNameLabel = 496
    xlDialogSeriesShape = 504
    xlDialogChartOptionsDataLabels = 505
    xlDialogChartOptionsDataTable = 506
    xlDialogSetBackgroundPicture = 509
    xlDialogDataValidation = 525
    xlDialogChartType = 526
    xlDialogChartLocation = 527
    xlDialogExternalDataProperties = 530
    _xlDialogPhonetic = 538
    xlDialogChartSourceData = 540
    _xlDialogChartSourceData = 541
    xlDialogSeriesOptions = 557
    xlDialogPivotTableOptions = 567
    xlDialogPivotSolveOrder = 568
    xlDialogPivotCalculatedField = 570
    xlDialogPivotCalculatedItem = 572
    xlDialogConditionalFormatting = 583
    xlDialogInsertHyperlink = 596
    xlDialogProtectSharing = 620
    xlDialogOptionsME = 647
    xlDialogPublishAsWebPage = 653
    xlDialogPhonetic = 656
    xlDialogImportTextFile = 666
    xlDialogNewWebQuery = 667
    xlDialogWebOptionsGeneral = 683
    xlDialogWebOptionsFiles = 684
    xlDialogWebOptionsPictures = 685
    xlDialogWebOptionsEncoding = 686
    xlDialogWebOptionsFonts = 687
    xlDialogPivotClientServerSet = 689
    xlDialogEvaluateFormula = 709
    xlDialogDataLabelMultiple = 723
    xlDialogChartOptionsDataLabelMultiple = 724
    xlDialogSearch = 731
    xlDialogErrorChecking = 732
    xlDialogPropertyFields = 754
    xlDialogWebOptionsBrowsers = 773
    xlDialogCreateList = 796
    xlDialogPermission = 832
    xlDialogMyPermission = 834
    xlDialogDocumentInspector = 862
    xlDialogNameManager = 977
    xlDialogNewName = 978
    xlDialogSetTupleEditorOnRows = 1107
    xlDialogSetTupleEditorOnColumns = 1108
    xlDialogSetManager = 1109
    xlDialogSparklineInsertLine = 1133
    xlDialogSparklineInsertColumn = 1134
    xlDialogSparklineInsertWinLoss = 1135
    xlDialogPivotTableWhatIfAnalysisSettings = 1153
    xlDialogSlicerSettings = 1179
    xlDialogSlicerCreation = 1182
    xlDialogPivotTableSlicerConnections = 1183
    xlDialogSlicerPivotTableConnections = 1184
    xlDialogSetMDXEditor = 1208
    xlDialogRecommendedPivotTables = 1258
    xlDialogManageRelationships = 1271
    xlDialogCreateRelationship = 1272
    xlDialogForecastETS = 1300
    xlDialogPivotDefaultLayout = 1360


class XlRoutingSlipDelivery(IntEnum):
    xlOneAfterAnother = 1
    xlAllAtOnce = 2


class XlPivotTableSourceType(IntEnum):
    xlPivotTable = -4148
    xlDatabase = 1
    xlExternal = 2
    xlConsolidation = 3
    xlScenario = 4


class XlCalculation(IntEnum):
    xlCalculationManual = -4135
    xlCalculationAutomatic = -4105
    xlCalculationSemiautomatic = 2


class XlCalcMemNumberFormatType(IntEnum):
    xlNumberFormatTypeDefault = 0
    xlNumberFormatTypeNumber = 1
    xlNumberFormatTypePercent = 2


class XlRoutingSlipStatus(IntEnum):
    xlNotYetRouted = 0
    xlRoutingInProgress = 1
    xlRoutingComplete = 2


class XlTimelineLevel(IntEnum):
    xlTimelineLevelYears = 0
    xlTimelineLevelQuarters = 1
    xlTimelineLevelMonths = 2
    xlTimelineLevelDays = 3


class XlPictureAppearance(IntEnum):
    xlScreen = 1
    xlPrinter = 2


class XlSaveConflictResolution(IntEnum):
    xlUserResolution = 1
    xlLocalSessionChanges = 2
    xlOtherSessionChanges = 3


class XlFileValidationPivotMode(IntEnum):
    xlFileValidationPivotDefault = 0
    xlFileValidationPivotRun = 1
    xlFileValidationPivotSkip = 2


class XlFormulaLabel(IntEnum):
    xlNoLabels = -4142
    xlRowLabels = 1
    xlColumnLabels = 2
    xlMixedLabels = 3


class XlChartLocation(IntEnum):
    xlLocationAsNewSheet = 1
    xlLocationAsObject = 2
    xlLocationAutomatic = 3


class XlEnableSelection(IntEnum):
    xlNoSelection = -4142
    xlNoRestrictions = 0
    xlUnlockedCells = 1


class XlRangeAutoFormat(IntEnum):
    xlRangeAutoFormatSimple = -4154
    xlRangeAutoFormatNone = -4142
    xlRangeAutoFormatClassic1 = 1
    xlRangeAutoFormatClassic2 = 2
    xlRangeAutoFormatClassic3 = 3
    xlRangeAutoFormatAccounting1 = 4
    xlRangeAutoFormatAccounting2 = 5
    xlRangeAutoFormatAccounting3 = 6
    xlRangeAutoFormatColor1 = 7
    xlRangeAutoFormatColor2 = 8
    xlRangeAutoFormatColor3 = 9
    xlRangeAutoFormatList1 = 10
    xlRangeAutoFormatList2 = 11
    xlRangeAutoFormatList3 = 12
    xlRangeAutoFormat3DEffects1 = 13
    xlRangeAutoFormat3DEffects2 = 14
    xlRangeAutoFormatLocalFormat1 = 15
    xlRangeAutoFormatLocalFormat2 = 16
    xlRangeAutoFormatAccounting4 = 17
    xlRangeAutoFormatLocalFormat3 = 19
    xlRangeAutoFormatLocalFormat4 = 20
    xlRangeAutoFormatReport1 = 21
    xlRangeAutoFormatReport2 = 22
    xlRangeAutoFormatReport3 = 23
    xlRangeAutoFormatReport4 = 24
    xlRangeAutoFormatReport5 = 25
    xlRangeAutoFormatReport6 = 26
    xlRangeAutoFormatReport7 = 27
    xlRangeAutoFormatReport8 = 28
    xlRangeAutoFormatReport9 = 29
    xlRangeAutoFormatReport10 = 30
    xlRangeAutoFormatClassicPivotTable = 31
    xlRangeAutoFormatTable1 = 32
    xlRangeAutoFormatTable2 = 33
    xlRangeAutoFormatTable3 = 34
    xlRangeAutoFormatTable4 = 35
    xlRangeAutoFormatTable5 = 36
    xlRangeAutoFormatTable6 = 37
    xlRangeAutoFormatTable7 = 38
    xlRangeAutoFormatTable8 = 39
    xlRangeAutoFormatTable9 = 40
    xlRangeAutoFormatTable10 = 41
    xlRangeAutoFormatPTNone = 42


class XlSortOn(IntEnum):
    xlSortOnValues = 0
    xlSortOnCellColor = 1
    xlSortOnFontColor = 2
    xlSortOnIcon = 3


class XlSparkScale(IntEnum):
    xlSparkScaleGroup = 1
    xlSparkScaleSingle = 2
    xlSparkScaleCustom = 3


class XlCellInsertionMode(IntEnum):
    xlOverwriteCells = 0
    xlInsertDeleteCells = 1
    xlInsertEntireRows = 2


class XlPivotLineType(IntEnum):
    xlPivotLineRegular = 0
    xlPivotLineSubtotal = 1
    xlPivotLineGrandTotal = 2
    xlPivotLineBlank = 3


class XlListDataType(IntEnum):
    xlListDataTypeNone = 0
    xlListDataTypeText = 1
    xlListDataTypeMultiLineText = 2
    xlListDataTypeNumber = 3
    xlListDataTypeCurrency = 4
    xlListDataTypeDateTime = 5
    xlListDataTypeChoice = 6
    xlListDataTypeChoiceMulti = 7
    xlListDataTypeListLookup = 8
    xlListDataTypeCheckbox = 9
    xlListDataTypeHyperLink = 10
    xlListDataTypeCounter = 11
    xlListDataTypeMultiLineRichText = 12


class XlEnableCancelKey(IntEnum):
    xlDisabled = 0
    xlInterrupt = 1
    xlErrorHandler = 2


class XlTickLabelOrientation(IntEnum):
    xlTickLabelOrientationUpward = -4171
    xlTickLabelOrientationDownward = -4170
    xlTickLabelOrientationVertical = -4166
    xlTickLabelOrientationHorizontal = -4128
    xlTickLabelOrientationAutomatic = -4105


class XlIconSet(IntEnum):
    xlCustomSet = -1
    xl3Arrows = 1
    xl3ArrowsGray = 2
    xl3Flags = 3
    xl3TrafficLights1 = 4
    xl3TrafficLights2 = 5
    xl3Signs = 6
    xl3Symbols = 7
    xl3Symbols2 = 8
    xl4Arrows = 9
    xl4ArrowsGray = 10
    xl4RedToBlack = 11
    xl4CRV = 12
    xl4TrafficLights = 13
    xl5Arrows = 14
    xl5ArrowsGray = 15
    xl5CRV = 16
    xl5Quarters = 17
    xl3Stars = 18
    xl3Triangles = 19
    xl5Boxes = 20


class XlXmlExportResult(IntEnum):
    xlXmlExportSuccess = 0
    xlXmlExportValidationFailed = 1


class XlArabicModes(IntEnum):
    xlArabicNone = 0
    xlArabicStrictAlefHamza = 1
    xlArabicStrictFinalYaa = 2
    xlArabicBothStrict = 3


class XlThemeFont(IntEnum):
    xlThemeFontNone = 0
    xlThemeFontMajor = 1
    xlThemeFontMinor = 2


class XlPlacement(IntEnum):
    xlMoveAndSize = 1
    xlMove = 2
    xlFreeFloating = 3


class XlSubtototalLocationType(IntEnum):
    xlAtTop = 1
    xlAtBottom = 2


class XlAllocation(IntEnum):
    xlManualAllocation = 1
    xlAutomaticAllocation = 2


class XlQueryType(IntEnum):
    xlODBCQuery = 1
    xlDAORecordset = 2
    xlWebQuery = 4
    xlOLEDBQuery = 5
    xlTextImport = 6
    xlADORecordset = 7


class XlSizeRepresents(IntEnum):
    xlSizeIsArea = 1
    xlSizeIsWidth = 2


class XlTableStyleElementType(IntEnum):
    xlWholeTable = 0
    xlHeaderRow = 1
    xlGrandTotalRow = 2
    xlTotalRow = 2
    xlFirstColumn = 3
    xlGrandTotalColumn = 4
    xlLastColumn = 4
    xlRowStripe1 = 5
    xlRowStripe2 = 6
    xlColumnStripe1 = 7
    xlColumnStripe2 = 8
    xlFirstHeaderCell = 9
    xlLastHeaderCell = 10
    xlFirstTotalCell = 11
    xlLastTotalCell = 12
    xlSubtotalColumn1 = 13
    xlSubtotalColumn2 = 14
    xlSubtotalColumn3 = 15
    xlSubtotalRow1 = 16
    xlSubtotalRow2 = 17
    xlSubtotalRow3 = 18
    xlBlankRow = 19
    xlColumnSubheading1 = 20
    xlColumnSubheading2 = 21
    xlColumnSubheading3 = 22
    xlRowSubheading1 = 23
    xlRowSubheading2 = 24
    xlRowSubheading3 = 25
    xlPageFieldLabels = 26
    xlPageFieldValues = 27
    xlSlicerUnselectedItemWithData = 28
    xlSlicerUnselectedItemWithNoData = 29
    xlSlicerSelectedItemWithData = 30
    xlSlicerSelectedItemWithNoData = 31
    xlSlicerHoveredUnselectedItemWithData = 32
    xlSlicerHoveredSelectedItemWithData = 33
    xlSlicerHoveredUnselectedItemWithNoData = 34
    xlSlicerHoveredSelectedItemWithNoData = 35
    xlTimelineSelectionLabel = 36
    xlTimelineTimeLevel = 37
    xlTimelinePeriodLabels1 = 38
    xlTimelinePeriodLabels2 = 39
    xlTimelineSelectedTimeBlock = 40
    xlTimelineUnselectedTimeBlock = 41
    xlTimelineSelectedTimeBlockSpace = 42


class XlDataSeriesDate(IntEnum):
    xlDay = 1
    xlWeekday = 2
    xlMonth = 3
    xlYear = 4


class XlPortugueseReform(IntEnum):
    xlPortuguesePreReform = 1
    xlPortuguesePostReform = 2
    xlPortugueseBoth = 3


class XlSeriesNameLevel(IntEnum):
    xlSeriesNameLevelNone = -3
    xlSeriesNameLevelCustom = -2
    xlSeriesNameLevelAll = -1


class XlEditionType(IntEnum):
    xlPublisher = 1
    xlSubscriber = 2


class XlPrintLocation(IntEnum):
    xlPrintNoComments = -4142
    xlPrintSheetEnd = 1
    xlPrintInPlace = 16


class XlVAlign(IntEnum):
    xlVAlignTop = -4160
    xlVAlignJustify = -4130
    xlVAlignDistributed = -4117
    xlVAlignCenter = -4108
    xlVAlignBottom = -4107


class XlXmlImportResult(IntEnum):
    xlXmlImportSuccess = 0
    xlXmlImportElementsTruncated = 1
    xlXmlImportValidationFailed = 2


class XlCalculatedMemberType(IntEnum):
    xlCalculatedMember = 0
    xlCalculatedSet = 1
    xlCalculatedMeasure = 2


class XlLegendPosition(IntEnum):
    xlLegendPositionCustom = -4161
    xlLegendPositionTop = -4160
    xlLegendPositionRight = -4152
    xlLegendPositionLeft = -4131
    xlLegendPositionBottom = -4107
    xlLegendPositionCorner = 2


class XlLinkType(IntEnum):
    xlLinkTypeExcelLinks = 1
    xlLinkTypeOLELinks = 2


class XlFormControl(IntEnum):
    xlButtonControl = 0
    xlCheckBox = 1
    xlDropDown = 2
    xlEditBox = 3
    xlGroupBox = 4
    xlLabel = 5
    xlListBox = 6
    xlOptionButton = 7
    xlScrollBar = 8
    xlSpinner = 9


class XlPageBreakExtent(IntEnum):
    xlPageBreakFull = 1
    xlPageBreakPartial = 2


class XlCellType(IntEnum):
    xlCellTypeSameValidation = -4175
    xlCellTypeAllValidation = -4174
    xlCellTypeSameFormatConditions = -4173
    xlCellTypeAllFormatConditions = -4172
    xlCellTypeComments = -4144
    xlCellTypeFormulas = -4123
    xlCellTypeConstants = 2
    xlCellTypeBlanks = 4
    xlCellTypeLastCell = 11
    xlCellTypeVisible = 12


class XlCommandUnderlines(IntEnum):
    xlCommandUnderlinesOff = -4146
    xlCommandUnderlinesAutomatic = -4105
    xlCommandUnderlinesOn = 1


class XlListConflict(IntEnum):
    xlListConflictDialog = 0
    xlListConflictRetryAllConflicts = 1
    xlListConflictDiscardAllConflicts = 2
    xlListConflictError = 3


class XlListObjectSourceType(IntEnum):
    xlSrcExternal = 0
    xlSrcRange = 1
    xlSrcXml = 2
    xlSrcQuery = 3
    xlSrcModel = 4


class XlPivotFormatType(IntEnum):
    xlReport1 = 0
    xlReport2 = 1
    xlReport3 = 2
    xlReport4 = 3
    xlReport5 = 4
    xlReport6 = 5
    xlReport7 = 6
    xlReport8 = 7
    xlReport9 = 8
    xlReport10 = 9
    xlTable1 = 10
    xlTable2 = 11
    xlTable3 = 12
    xlTable4 = 13
    xlTable5 = 14
    xlTable6 = 15
    xlTable7 = 16
    xlTable8 = 17
    xlTable9 = 18
    xlTable10 = 19
    xlPTClassic = 20
    xlPTNone = 21


class XlCutCopyMode(IntEnum):
    xlCopy = 1
    xlCut = 2


class XlPivotFieldOrientation(IntEnum):
    xlHidden = 0
    xlRowField = 1
    xlColumnField = 2
    xlPageField = 3
    xlDataField = 4


class XlFixedFormatType(IntEnum):
    xlTypePDF = 0
    xlTypeXPS = 1


class XlUpdateLinks(IntEnum):
    xlUpdateLinksUserSetting = 1
    xlUpdateLinksNever = 2
    xlUpdateLinksAlways = 3


class XlSparklineRowCol(IntEnum):
    xlSparklineNonSquare = 0
    xlSparklineRowsSquare = 1
    xlSparklineColumnsSquare = 2


class XlYesNoGuess(IntEnum):
    xlGuess = 0
    xlYes = 1
    xlNo = 2


class XlChartElementPosition(IntEnum):
    xlChartElementPositionCustom = -4114
    xlChartElementPositionAutomatic = -4105


class XlSortMethod(IntEnum):
    xlPinYin = 1
    xlStroke = 2


class XlSortOrder(IntEnum):
    xlAscending = 1
    xlDescending = 2


class XlMSApplication(IntEnum):
    xlMicrosoftWord = 1
    xlMicrosoftPowerPoint = 2
    xlMicrosoftMail = 3
    xlMicrosoftAccess = 4
    xlMicrosoftFoxPro = 5
    xlMicrosoftProject = 6
    xlMicrosoftSchedulePlus = 7


class XlDirection(IntEnum):
    xlUp = -4162
    xlToRight = -4161
    xlToLeft = -4159
    xlDown = -4121


class XlCmdType(IntEnum):
    xlCmdCube = 1
    xlCmdSql = 2
    xlCmdTable = 3
    xlCmdDefault = 4
    xlCmdList = 5
    xlCmdTableCollection = 6
    xlCmdExcel = 7
    xlCmdDAX = 8


class XlGenerateTableRefs(IntEnum):
    xlGenerateTableRefA1 = 0
    xlGenerateTableRefStruct = 1


class XlFileAccess(IntEnum):
    xlReadWrite = 2
    xlReadOnly = 3


class XlPivotFieldRepeatLabels(IntEnum):
    xlDoNotRepeatLabels = 1
    xlRepeatLabels = 2


class XlAutoFilterOperator(IntEnum):
    xlAnd = 1
    xlOr = 2
    xlTop10Items = 3
    xlBottom10Items = 4
    xlTop10Percent = 5
    xlBottom10Percent = 6
    xlFilterValues = 7
    xlFilterCellColor = 8
    xlFilterFontColor = 9
    xlFilterIcon = 10
    xlFilterDynamic = 11
    xlFilterNoFill = 12
    xlFilterAutomaticFontColor = 13
    xlFilterNoIcon = 14


class XlPlatform(IntEnum):
    xlMacintosh = 1
    xlWindows = 2
    xlMSDOS = 3


class XlSubscribeToFormat(IntEnum):
    xlSubscribeToText = -4158
    xlSubscribeToPicture = -4147


class XlRemoveDocInfoType(IntEnum):
    xlRDIComments = 1
    xlRDIRemovePersonalInformation = 4
    xlRDIEmailHeader = 5
    xlRDIRoutingSlip = 6
    xlRDISendForReview = 7
    xlRDIDocumentProperties = 8
    xlRDIDocumentWorkspace = 10
    xlRDIInkAnnotations = 11
    xlRDIScenarioComments = 12
    xlRDIPublishInfo = 13
    xlRDIDocumentServerProperties = 14
    xlRDIDocumentManagementPolicy = 15
    xlRDIContentType = 16
    xlRDIDefinedNameComments = 18
    xlRDIInactiveDataConnections = 19
    xlRDIPrinterPath = 20
    xlRDIInlineWebExtensions = 21
    xlRDITaskpaneWebExtensions = 22
    xlRDIExcelDataModel = 23
    xlRDIAll = 99


class XlWindowState(IntEnum):
    xlNormal = -4143
    xlMinimized = -4140
    xlMaximized = -4137


class XlSummaryRow(IntEnum):
    xlSummaryAbove = 0
    xlSummaryBelow = 1


class XlAutoFillType(IntEnum):
    xlFillDefault = 0
    xlFillCopy = 1
    xlFillSeries = 2
    xlFillFormats = 3
    xlFillValues = 4
    xlFillDays = 5
    xlFillWeekdays = 6
    xlFillMonths = 7
    xlFillYears = 8
    xlLinearTrend = 9
    xlGrowthTrend = 10
    xlFlashFill = 11


class XlFormatConditionType(IntEnum):
    xlCellValue = 1
    xlExpression = 2
    xlColorScale = 3
    xlDatabar = 4
    xlTop10 = 5
    xlIconSets = 6
    xlUniqueValues = 8
    xlTextString = 9
    xlBlanksCondition = 10
    xlTimePeriod = 11
    xlAboveAverageCondition = 12
    xlNoBlanksCondition = 13
    xlErrorsCondition = 16
    xlNoErrorsCondition = 17


class XlCalculationInterruptKey(IntEnum):
    xlNoKey = 0
    xlEscKey = 1
    xlAnyKey = 2


class XlHebrewModes(IntEnum):
    xlHebrewFullScript = 0
    xlHebrewPartialScript = 1
    xlHebrewMixedScript = 2
    xlHebrewMixedAuthorizedScript = 3


class XlXLMMacroType(IntEnum):
    xlFunction = 1
    xlCommand = 2
    xlNotXLM = 3


class XlPivotCellType(IntEnum):
    xlPivotCellValue = 0
    xlPivotCellPivotItem = 1
    xlPivotCellSubtotal = 2
    xlPivotCellGrandTotal = 3
    xlPivotCellDataField = 4
    xlPivotCellPivotField = 5
    xlPivotCellPageFieldItem = 6
    xlPivotCellCustomSubtotal = 7
    xlPivotCellDataPivotField = 8
    xlPivotCellBlankCell = 9


class XlPivotTableVersionList(IntEnum):
    xlPivotTableVersionCurrent = -1
    xlPivotTableVersion2000 = 0
    xlPivotTableVersion10 = 1
    xlPivotTableVersion11 = 2
    xlPivotTableVersion12 = 3
    xlPivotTableVersion14 = 4
    xlPivotTableVersion15 = 5


class XlSlicerCacheType(IntEnum):
    xlSlicer = 1
    xlTimeline = 2


class XlPublishToDocsDisclosureScope(IntEnum):
    msoPublic = 0
    msoLimited = 1
    msoOrganization = 2
    msoNoOverwrite = 3


class XlDataSeriesType(IntEnum):
    xlDataSeriesLinear = -4132
    xlGrowth = 2
    xlChronological = 3
    xlAutoFill = 4


class XlLinkInfo(IntEnum):
    xlUpdateState = 1
    xlEditionDate = 2
    xlLinkInfoStatus = 3


class XlCopyPictureFormat(IntEnum):
    xlPicture = -4147
    xlBitmap = 2


class XlSpanishModes(IntEnum):
    xlSpanishTuteoOnly = 0
    xlSpanishTuteoAndVoseo = 1
    xlSpanishVoseoOnly = 2

