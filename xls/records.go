package xls

import "fmt"

type recordType uint16

// Record types defined by the XLS specification document, section 2.3/2.4.
// https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/43684742-8fcd-4fcd-92df-157d8d7241f9
const (
	RecTypeFormula              recordType = 6    // per section 2.4.127
	RecTypeEOF                  recordType = 10   // section 2.4.103
	RecTypeCalcCount            recordType = 12   // section 2.4.31
	RecTypeCalcMode             recordType = 13   // section 2.4.34
	RecTypeCalcPrecision        recordType = 14   // section 2.4.35
	RecTypeCalcRefMode          recordType = 15   // section 2.4.36
	RecTypeCalcDelta            recordType = 16   // section 2.4.32
	RecTypeCalcIter             recordType = 17   // section 2.4.33
	RecTypeProtect              recordType = 18   // section 2.4.207
	RecTypePassword             recordType = 19   // section 2.4.191
	RecTypeHeader               recordType = 20   // section 2.4.136
	RecTypeFooter               recordType = 21   // section 2.4.124
	RecTypeExternSheet          recordType = 23   // section 2.4.106
	RecTypeLbl                  recordType = 24   // section 2.4.150
	RecTypeWinProtect           recordType = 25   // section 2.4.347
	RecTypeVerticalPageBreaks   recordType = 26   // section 2.4.343
	RecTypeHorizontalPageBreaks recordType = 27   // section 2.4.142
	RecTypeNote                 recordType = 28   // section 2.4.179
	RecTypeSelection            recordType = 29   // section 2.4.248
	RecTypeDate1904             recordType = 34   // section 2.4.77
	RecTypeExternName           recordType = 35   // section 2.4.105
	RecTypeLeftMargin           recordType = 38   // section 2.4.151
	RecTypeRightMargin          recordType = 39   // section 2.4.219
	RecTypeTopMargin            recordType = 40   // section 2.4.328
	RecTypeBottomMargin         recordType = 41   // section 2.4.27
	RecTypePrintRowCol          recordType = 42   // section 2.4.203
	RecTypePrintGrid            recordType = 43   // section 2.4.202
	RecTypeFilePass             recordType = 47   // section 2.4.117
	RecTypeFont                 recordType = 49   // section 2.4.122
	RecTypePrintSize            recordType = 51   // section 2.4.204
	RecTypeContinue             recordType = 60   // section 2.4.58
	RecTypeWindow1              recordType = 61   // section 2.4.345
	RecTypeBackup               recordType = 64   // section 2.4.14
	RecTypePane                 recordType = 65   // section 2.4.189
	RecTypeCodePage             recordType = 66   // section 2.4.52
	RecTypePls                  recordType = 77   // section 2.4.199
	RecTypeDCon                 recordType = 80   // section 2.4.82
	RecTypeDConRef              recordType = 81   // section 2.4.86
	RecTypeDConName             recordType = 82   // section 2.4.85
	RecTypeDefColWidth          recordType = 85   // section 2.4.89
	RecTypeXCT                  recordType = 89   // section 2.4.352
	RecTypeCRN                  recordType = 90   // section 2.4.65
	RecTypeFileSharing          recordType = 91   // section 2.4.118
	RecTypeWriteAccess          recordType = 92   // section 2.4.349
	RecTypeObj                  recordType = 93   // section 2.4.181
	RecTypeUncalced             recordType = 94   // section 2.4.331
	RecTypeCalcSaveRecalc       recordType = 95   // section 2.4.37
	RecTypeTemplate             recordType = 96   // section 2.4.323
	RecTypeIntl                 recordType = 97   // section 2.4.147
	RecTypeObjProtect           recordType = 99   // section 2.4.183
	RecTypeColInfo              recordType = 125  // section 2.4.53
	RecTypeGuts                 recordType = 128  // section 2.4.134
	RecTypeWsBool               recordType = 129  // section 2.4.351
	RecTypeGridSet              recordType = 130  // section 2.4.132
	RecTypeHCenter              recordType = 131  // section 2.4.135
	RecTypeVCenter              recordType = 132  // section 2.4.342
	RecTypeBoundSheet8          recordType = 133  // section 2.4.28
	RecTypeWriteProtect         recordType = 134  // section 2.4.350
	RecTypeCountry              recordType = 140  // section 2.4.63
	RecTypeHideObj              recordType = 141  // section 2.4.139
	RecTypeSort                 recordType = 144  // section 2.4.263
	RecTypePalette              recordType = 146  // section 2.4.188
	RecTypeSync                 recordType = 151  // section 2.4.318
	RecTypeLPr                  recordType = 152  // section 2.4.158
	RecTypeDxGCol               recordType = 153  // section 2.4.98
	RecTypeFnGroupName          recordType = 154  // section 2.4.120
	RecTypeFilterMode           recordType = 155  // section 2.4.119
	RecTypeBuiltInFnGroupCount  recordType = 156  // section 2.4.30
	RecTypeAutoFilterInfo       recordType = 157  // section 2.4.8
	RecTypeAutoFilter           recordType = 158  // section 2.4.6
	RecTypeScl                  recordType = 160  // section 2.4.247
	RecTypeSetup                recordType = 161  // section 2.4.257
	RecTypeScenMan              recordType = 174  // section 2.4.246
	RecTypeSCENARIO             recordType = 175  // section 2.4.244
	RecTypeSxView               recordType = 176  // section 2.4.313
	RecTypeSxvd                 recordType = 177  // section 2.4.309
	RecTypeSXVI                 recordType = 178  // section 2.4.312
	RecTypeSxIvd                recordType = 180  // section 2.4.292
	RecTypeSXLI                 recordType = 181  // section 2.4.293
	RecTypeSXPI                 recordType = 182  // section 2.4.298
	RecTypeDocRoute             recordType = 184  // section 2.4.91
	RecTypeRecipName            recordType = 185  // section 2.4.216
	RecTypeMulRk                recordType = 189  // section 2.4.175
	RecTypeMulBlank             recordType = 190  // section 2.4.174
	RecTypeMms                  recordType = 193  // section 2.4.169
	RecTypeSXDI                 recordType = 197  // section 2.4.278
	RecTypeSXDB                 recordType = 198  // section 2.4.275
	RecTypeSXFDB                recordType = 199  // section 2.4.283
	RecTypeSXDBB                recordType = 200  // section 2.4.276
	RecTypeSXNum                recordType = 201  // section 2.4.296
	RecTypeSxBool               recordType = 202  // section 2.4.274
	RecTypeSxErr                recordType = 203  // section 2.4.281
	RecTypeSXInt                recordType = 204  // section 2.4.289
	RecTypeSXString             recordType = 205  // section 2.4.304
	RecTypeSXDtr                recordType = 206  // section 2.4.279
	RecTypeSxNil                recordType = 207  // section 2.4.295
	RecTypeSXTbl                recordType = 208  // section 2.4.305
	RecTypeSXTBRGIITM           recordType = 209  // section 2.4.307
	RecTypeSxTbpg               recordType = 210  // section 2.4.306
	RecTypeObProj               recordType = 211  // section 2.4.185
	RecTypeSXStreamID           recordType = 213  // section 2.4.303
	RecTypeDBCell               recordType = 215  // section 2.4.78
	RecTypeSXRng                recordType = 216  // section 2.4.300
	RecTypeSxIsxoper            recordType = 217  // section 2.4.290
	RecTypeBookBool             recordType = 218  // section 2.4.22
	RecTypeDbOrParamQry         recordType = 220  // section 2.4.79
	RecTypeScenarioProtect      recordType = 221  // section 2.4.245
	RecTypeOleObjectSize        recordType = 222  // section 2.4.187
	RecTypeXF                   recordType = 224  // section 2.4.353
	RecTypeInterfaceHdr         recordType = 225  // section 2.4.146
	RecTypeInterfaceEnd         recordType = 226  // section 2.4.145
	RecTypeSXVS                 recordType = 227  // section 2.4.317
	RecTypeMergeCells           recordType = 229  // section 2.4.168
	RecTypeBkHim                recordType = 233  // section 2.4.19
	RecTypeMsoDrawingGroup      recordType = 235  // section 2.4.171
	RecTypeMsoDrawing           recordType = 236  // section 2.4.170
	RecTypeMsoDrawingSelection  recordType = 237  // section 2.4.172
	RecTypePhoneticInfo         recordType = 239  // section 2.4.192
	RecTypeSxRule               recordType = 240  // section 2.4.301
	RecTypeSXEx                 recordType = 241  // section 2.4.282
	RecTypeSxFilt               recordType = 242  // section 2.4.285
	RecTypeSxDXF                recordType = 244  // section 2.4.280
	RecTypeSxItm                recordType = 245  // section 2.4.291
	RecTypeSxName               recordType = 246  // section 2.4.294
	RecTypeSxSelect             recordType = 247  // section 2.4.302
	RecTypeSXPair               recordType = 248  // section 2.4.297
	RecTypeSxFmla               recordType = 249  // section 2.4.286
	RecTypeSxFormat             recordType = 251  // section 2.4.287
	RecTypeSST                  recordType = 252  // section 2.4.265
	RecTypeLabelSst             recordType = 253  // section 2.4.149
	RecTypeExtSST               recordType = 255  // section 2.4.107
	RecTypeSXVDEx               recordType = 256  // section 2.4.310
	RecTypeSXFormula            recordType = 259  // section 2.4.288
	RecTypeSXDBEx               recordType = 290  // section 2.4.277
	RecTypeRRDInsDel            recordType = 311  // section 2.4.228
	RecTypeRRDHead              recordType = 312  // section 2.4.226
	RecTypeRRDChgCell           recordType = 315  // section 2.4.223
	RecTypeRRTabID              recordType = 317  // section 2.4.241
	RecTypeRRDRenSheet          recordType = 318  // section 2.4.234
	RecTypeRRSort               recordType = 319  // section 2.4.240
	RecTypeRRDMove              recordType = 320  // section 2.4.231
	RecTypeRRFormat             recordType = 330  // section 2.4.238
	RecTypeRRAutoFmt            recordType = 331  // section 2.4.222
	RecTypeRRInsertSh           recordType = 333  // section 2.4.239
	RecTypeRRDMoveBegin         recordType = 334  // section 2.4.232
	RecTypeRRDMoveEnd           recordType = 335  // section 2.4.233
	RecTypeRRDInsDelBegin       recordType = 336  // section 2.4.229
	RecTypeRRDInsDelEnd         recordType = 337  // section 2.4.230
	RecTypeRRDConflict          recordType = 338  // section 2.4.224
	RecTypeRRDDefName           recordType = 339  // section 2.4.225
	RecTypeRRDRstEtxp           recordType = 340  // section 2.4.235
	RecTypeLRng                 recordType = 351  // section 2.4.159
	RecTypeUsesELFs             recordType = 352  // section 2.4.337
	RecTypeDSF                  recordType = 353  // section 2.4.94
	RecTypeCUsr                 recordType = 401  // section 2.4.72
	RecTypeCbUsr                recordType = 402  // section 2.4.40
	RecTypeUsrInfo              recordType = 403  // section 2.4.340
	RecTypeUsrExcl              recordType = 404  // section 2.4.339
	RecTypeFileLock             recordType = 405  // section 2.4.116
	RecTypeRRDInfo              recordType = 406  // section 2.4.227
	RecTypeBCUsrs               recordType = 407  // section 2.4.16
	RecTypeUsrChk               recordType = 408  // section 2.4.338
	RecTypeUserBView            recordType = 425  // section 2.4.333
	RecTypeUserSViewBegin       recordType = 426  // section 2.4.334
	RecTypeUserSViewBeginChart  recordType = 426  // section 2.4.335
	RecTypeUserSViewEnd         recordType = 427  // section 2.4.336
	RecTypeRRDUserView          recordType = 428  // section 2.4.237
	RecTypeQsi                  recordType = 429  // section 2.4.208
	RecTypeSupBook              recordType = 430  // section 2.4.271
	RecTypeProt4Rev             recordType = 431  // section 2.4.205
	RecTypeCondFmt              recordType = 432  // section 2.4.56
	RecTypeCF                   recordType = 433  // section 2.4.42
	RecTypeDVal                 recordType = 434  // section 2.4.96
	RecTypeDConBin              recordType = 437  // section 2.4.83
	RecTypeTxO                  recordType = 438  // section 2.4.329
	RecTypeRefreshAll           recordType = 439  // section 2.4.217
	RecTypeHLink                recordType = 440  // section 2.4.140
	RecTypeLel                  recordType = 441  // section 2.4.154
	RecTypeCodeName             recordType = 442  // section 2.4.51
	RecTypeSXFDBType            recordType = 443  // section 2.4.284
	RecTypeProt4RevPass         recordType = 444  // section 2.4.206
	RecTypeObNoMacros           recordType = 445  // section 2.4.184
	RecTypeDv                   recordType = 446  // section 2.4.95
	RecTypeExcel9File           recordType = 448  // section 2.4.104
	RecTypeRecalcID             recordType = 449  // section 2.4.215
	RecTypeEntExU2              recordType = 450  // section 2.4.102
	RecTypeDimensions           recordType = 512  // section 2.4.90
	RecTypeBlank                recordType = 513  // section 2.4.20
	RecTypeNumber               recordType = 515  // section 2.4.180
	RecTypeLabel                recordType = 516  // section 2.4.148
	RecTypeBoolErr              recordType = 517  // section 2.4.24
	RecTypeString               recordType = 519  // section 2.4.268
	RecTypeRow                  recordType = 520  // section 2.4.221
	RecTypeIndex                recordType = 523  // section 2.4.144
	RecTypeArray                recordType = 545  // section 2.4.4
	RecTypeDefaultRowHeight     recordType = 549  // section 2.4.87
	RecTypeTable                recordType = 566  // section 2.4.319
	RecTypeWindow2              recordType = 574  // section 2.4.346
	RecTypeRK                   recordType = 638  // section 2.4.220
	RecTypeStyle                recordType = 659  // section 2.4.269
	RecTypeBigName              recordType = 1048 // section 2.4.18
	RecTypeFormat               recordType = 1054 // section 2.4.126
	RecTypeContinueBigName      recordType = 1084 // section 2.4.59
	RecTypeShrFmla              recordType = 1212 // section 2.4.260
	RecTypeHLinkTooltip         recordType = 2048 // section 2.4.141
	RecTypeWebPub               recordType = 2049 // section 2.4.344
	RecTypeQsiSXTag             recordType = 2050 // section 2.4.211
	RecTypeDBQueryExt           recordType = 2051 // section 2.4.81
	RecTypeExtString            recordType = 2052 // section 2.4.108
	RecTypeTxtQry               recordType = 2053 // section 2.4.330
	RecTypeQsir                 recordType = 2054 // section 2.4.210
	RecTypeQsif                 recordType = 2055 // section 2.4.209
	RecTypeRRDTQSIF             recordType = 2056 // section 2.4.236
	RecTypeBOF                  recordType = 2057 // section 2.4.21
	RecTypeOleDbConn            recordType = 2058 // section 2.4.186
	RecTypeWOpt                 recordType = 2059 // section 2.4.348
	RecTypeSXViewEx             recordType = 2060 // section 2.4.314
	RecTypeSXTH                 recordType = 2061 // section 2.4.308
	RecTypeSXPIEx               recordType = 2062 // section 2.4.299
	RecTypeSXVDTEx              recordType = 2063 // section 2.4.311
	RecTypeSXViewEx9            recordType = 2064 // section 2.4.315
	RecTypeContinueFrt          recordType = 2066 // section 2.4.60
	RecTypeRealTimeData         recordType = 2067 // section 2.4.214
	RecTypeChartFrtInfo         recordType = 2128 // section 2.4.49
	RecTypeFrtWrapper           recordType = 2129 // section 2.4.130
	RecTypeStartBlock           recordType = 2130 // section 2.4.266
	RecTypeEndBlock             recordType = 2131 // section 2.4.100
	RecTypeStartObject          recordType = 2132 // section 2.4.267
	RecTypeEndObject            recordType = 2133 // section 2.4.101
	RecTypeCatLab               recordType = 2134 // section 2.4.38
	RecTypeYMult                recordType = 2135 // section 2.4.356
	RecTypeSXViewLink           recordType = 2136 // section 2.4.316
	RecTypePivotChartBits       recordType = 2137 // section 2.4.196
	RecTypeFrtFontList          recordType = 2138 // section 2.4.129
	RecTypeSheetExt             recordType = 2146 // section 2.4.259
	RecTypeBookExt              recordType = 2147 // section 2.4.23
	RecTypeSXAddl               recordType = 2148 // section 2.4.273.2
	RecTypeCrErr                recordType = 2149 // section 2.4.64
	RecTypeHFPicture            recordType = 2150 // section 2.4.138
	RecTypeFeatHdr              recordType = 2151 // section 2.4.112
	RecTypeFeat                 recordType = 2152 // section 2.4.111
	RecTypeDataLabExt           recordType = 2154 // section 2.4.75
	RecTypeDataLabExtContents   recordType = 2155 // section 2.4.76
	RecTypeCellWatch            recordType = 2156 // section 2.4.41
	RecTypeFeatHdr11            recordType = 2161 // section 2.4.113
	RecTypeFeature11            recordType = 2162 // section 2.4.114
	RecTypeDropDownObjIds       recordType = 2164 // section 2.4.93
	RecTypeContinueFrt11        recordType = 2165 // section 2.4.61
	RecTypeDConn                recordType = 2166 // section 2.4.84
	RecTypeList12               recordType = 2167 // section 2.4.157
	RecTypeFeature12            recordType = 2168 // section 2.4.115
	RecTypeCondFmt12            recordType = 2169 // section 2.4.57
	RecTypeCF12                 recordType = 2170 // section 2.4.43
	RecTypeCFEx                 recordType = 2171 // section 2.4.44
	RecTypeXFCRC                recordType = 2172 // section 2.4.354
	RecTypeXFExt                recordType = 2173 // section 2.4.355
	RecTypeAutoFilter12         recordType = 2174 // section 2.4.7
	RecTypeContinueFrt12        recordType = 2175 // section 2.4.62
	RecTypeMDTInfo              recordType = 2180 // section 2.4.162
	RecTypeMDXStr               recordType = 2181 // section 2.4.166
	RecTypeMDXTuple             recordType = 2182 // section 2.4.167
	RecTypeMDXSet               recordType = 2183 // section 2.4.165
	RecTypeMDXProp              recordType = 2184 // section 2.4.164
	RecTypeMDXKPI               recordType = 2185 // section 2.4.163
	RecTypeMDB                  recordType = 2186 // section 2.4.161
	RecTypePLV                  recordType = 2187 // section 2.4.200
	RecTypeCompat12             recordType = 2188 // section 2.4.54
	RecTypeDXF                  recordType = 2189 // section 2.4.97
	RecTypeTableStyles          recordType = 2190 // section 2.4.322
	RecTypeTableStyle           recordType = 2191 // section 2.4.320
	RecTypeTableStyleElement    recordType = 2192 // section 2.4.321
	RecTypeStyleExt             recordType = 2194 // section 2.4.270
	RecTypeNamePublish          recordType = 2195 // section 2.4.178
	RecTypeNameCmt              recordType = 2196 // section 2.4.176
	RecTypeSortData             recordType = 2197 // section 2.4.264
	RecTypeTheme                recordType = 2198 // section 2.4.326
	RecTypeGUIDTypeLib          recordType = 2199 // section 2.4.133
	RecTypeFnGrp12              recordType = 2200 // section 2.4.121
	RecTypeNameFnGrp12          recordType = 2201 // section 2.4.177
	RecTypeMTRSettings          recordType = 2202 // section 2.4.173
	RecTypeCompressPictures     recordType = 2203 // section 2.4.55
	RecTypeHeaderFooter         recordType = 2204 // section 2.4.137
	RecTypeCrtLayout12          recordType = 2205 // section 2.4.66
	RecTypeCrtMlFrt             recordType = 2206 // section 2.4.70
	RecTypeCrtMlFrtContinue     recordType = 2207 // section 2.4.71
	RecTypeForceFullCalculation recordType = 2211 // section 2.4.125
	RecTypeShapePropsStream     recordType = 2212 // section 2.4.258
	RecTypeTextPropsStream      recordType = 2213 // section 2.4.325
	RecTypeRichTextStream       recordType = 2214 // section 2.4.218
	RecTypeCrtLayout12A         recordType = 2215 // section 2.4.67
	RecTypeUnits                recordType = 4097 // section 2.4.332
	RecTypeChart                recordType = 4098 // section 2.4.45
	RecTypeSeries               recordType = 4099 // section 2.4.252
	RecTypeDataFormat           recordType = 4102 // section 2.4.74
	RecTypeLineFormat           recordType = 4103 // section 2.4.156
	RecTypeMarkerFormat         recordType = 4105 // section 2.4.160
	RecTypeAreaFormat           recordType = 4106 // section 2.4.3
	RecTypePieFormat            recordType = 4107 // section 2.4.195
	RecTypeAttachedLabel        recordType = 4108 // section 2.4.5
	RecTypeSeriesText           recordType = 4109 // section 2.4.254
	RecTypeChartFormat          recordType = 4116 // section 2.4.48
	RecTypeLegend               recordType = 4117 // section 2.4.152
	RecTypeSeriesList           recordType = 4118 // section 2.4.253
	RecTypeBar                  recordType = 4119 // section 2.4.15
	RecTypeLine                 recordType = 4120 // section 2.4.155
	RecTypePie                  recordType = 4121 // section 2.4.194
	RecTypeArea                 recordType = 4122 // section 2.4.2
	RecTypeScatter              recordType = 4123 // section 2.4.243
	RecTypeCrtLine              recordType = 4124 // section 2.4.68
	RecTypeAxis                 recordType = 4125 // section 2.4.11
	RecTypeTick                 recordType = 4126 // section 2.4.327
	RecTypeValueRange           recordType = 4127 // section 2.4.341
	RecTypeCatSerRange          recordType = 4128 // section 2.4.39
	RecTypeAxisLine             recordType = 4129 // section 2.4.12
	RecTypeCrtLink              recordType = 4130 // section 2.4.69
	RecTypeDefaultText          recordType = 4132 // section 2.4.88
	RecTypeText                 recordType = 4133 // section 2.4.324
	RecTypeFontX                recordType = 4134 // section 2.4.123
	RecTypeObjectLink           recordType = 4135 // section 2.4.182
	RecTypeFrame                recordType = 4146 // section 2.4.128
	RecTypeBegin                recordType = 4147 // section 2.4.17
	RecTypeEnd                  recordType = 4148 // section 2.4.99
	RecTypePlotArea             recordType = 4149 // section 2.4.197
	RecTypeChart3d              recordType = 4154 // section 2.4.46
	RecTypePicF                 recordType = 4156 // section 2.4.193
	RecTypeDropBar              recordType = 4157 // section 2.4.92
	RecTypeRadar                recordType = 4158 // section 2.4.212
	RecTypeSurf                 recordType = 4159 // section 2.4.272
	RecTypeRadarArea            recordType = 4160 // section 2.4.213
	RecTypeAxisParent           recordType = 4161 // section 2.4.13
	RecTypeLegendException      recordType = 4163 // section 2.4.153(
	RecTypeShtProps             recordType = 4164 // section 2.4.261
	RecTypeSerToCrt             recordType = 4165 // section 2.4.256
	RecTypeAxesUsed             recordType = 4166 // section 2.4.10
	RecTypeSBaseRef             recordType = 4168 // section 2.4.242
	RecTypeSerParent            recordType = 4170 // section 2.4.255
	RecTypeSerAuxTrend          recordType = 4171 // section 2.4.250
	RecTypeIFmtRecord           recordType = 4174 // section 2.4.143
	RecTypePos                  recordType = 4175 // section 2.4.201
	RecTypeAlRuns               recordType = 4176 // section 2.4.1
	RecTypeBRAI                 recordType = 4177 // section 2.4.29
	RecTypeSerAuxErrBar         recordType = 4187 // section 2.4.249
	RecTypeClrtClient           recordType = 4188 // section 2.4.50
	RecTypeSerFmt               recordType = 4189 // section 2.4.251
	RecTypeChart3DBarShape      recordType = 4191 // section 2.4.47
	RecTypeFbi                  recordType = 4192 // section 2.4.109
	RecTypeBopPop               recordType = 4193 // section 2.4.25
	RecTypeAxcExt               recordType = 4194 // section 2.4.9
	RecTypeDat                  recordType = 4195 // section 2.4.73
	RecTypePlotGrowth           recordType = 4196 // section 2.4.198
	RecTypeSIIndex              recordType = 4197 // section 2.4.262
	RecTypeGelFrame             recordType = 4198 // section 2.4.131
	RecTypeBopPopCustom         recordType = 4199 // section 2.4.26
	RecTypeFbi2                 recordType = 4200 // section 2.4.110
)

func (r recordType) String() string {
	switch r {
	case RecTypeFormula:
		return "Formula (6)"
	case RecTypeEOF:
		return "EOF (10)"
	case RecTypeCalcCount:
		return "CalcCount (12)"
	case RecTypeCalcMode:
		return "CalcMode (13)"
	case RecTypeCalcPrecision:
		return "CalcPrecision (14)"
	case RecTypeCalcRefMode:
		return "CalcRefMode (15)"
	case RecTypeCalcDelta:
		return "CalcDelta (16)"
	case RecTypeCalcIter:
		return "CalcIter (17)"
	case RecTypeProtect:
		return "Protect (18)"
	case RecTypePassword:
		return "Password (19)"
	case RecTypeHeader:
		return "Header (20)"
	case RecTypeFooter:
		return "Footer (21)"
	case RecTypeExternSheet:
		return "ExternSheet (23)"
	case RecTypeLbl:
		return "Lbl (24)"
	case RecTypeWinProtect:
		return "WinProtect (25)"
	case RecTypeVerticalPageBreaks:
		return "VerticalPageBreaks (26)"
	case RecTypeHorizontalPageBreaks:
		return "HorizontalPageBreaks (27)"
	case RecTypeNote:
		return "Note (28)"
	case RecTypeSelection:
		return "Selection (29)"
	case RecTypeDate1904:
		return "Date1904 (34)"
	case RecTypeExternName:
		return "ExternName (35)"
	case RecTypeLeftMargin:
		return "LeftMargin (38)"
	case RecTypeRightMargin:
		return "RightMargin (39)"
	case RecTypeTopMargin:
		return "TopMargin (40)"
	case RecTypeBottomMargin:
		return "BottomMargin (41)"
	case RecTypePrintRowCol:
		return "PrintRowCol (42)"
	case RecTypePrintGrid:
		return "PrintGrid (43)"
	case RecTypeFilePass:
		return "FilePass (47)"
	case RecTypeFont:
		return "Font (49)"
	case RecTypePrintSize:
		return "PrintSize (51)"
	case RecTypeContinue:
		return "Continue (60)"
	case RecTypeWindow1:
		return "Window1 (61)"
	case RecTypeBackup:
		return "Backup (64)"
	case RecTypePane:
		return "Pane (65)"
	case RecTypeCodePage:
		return "CodePage (66)"
	case RecTypePls:
		return "Pls (77)"
	case RecTypeDCon:
		return "DCon (80)"
	case RecTypeDConRef:
		return "DConRef (81)"
	case RecTypeDConName:
		return "DConName (82)"
	case RecTypeDefColWidth:
		return "DefColWidth (85)"
	case RecTypeXCT:
		return "XCT (89)"
	case RecTypeCRN:
		return "CRN (90)"
	case RecTypeFileSharing:
		return "FileSharing (91)"
	case RecTypeWriteAccess:
		return "WriteAccess (92)"
	case RecTypeObj:
		return "Obj (93)"
	case RecTypeUncalced:
		return "Uncalced (94)"
	case RecTypeCalcSaveRecalc:
		return "CalcSaveRecalc (95)"
	case RecTypeTemplate:
		return "Template (96)"
	case RecTypeIntl:
		return "Intl (97)"
	case RecTypeObjProtect:
		return "ObjProtect (99)"
	case RecTypeColInfo:
		return "ColInfo (125)"
	case RecTypeGuts:
		return "Guts (128)"
	case RecTypeWsBool:
		return "WsBool (129)"
	case RecTypeGridSet:
		return "GridSet (130)"
	case RecTypeHCenter:
		return "HCenter (131)"
	case RecTypeVCenter:
		return "VCenter (132)"
	case RecTypeBoundSheet8:
		return "BoundSheet8 (133)"
	case RecTypeWriteProtect:
		return "WriteProtect (134)"
	case RecTypeCountry:
		return "Country (140)"
	case RecTypeHideObj:
		return "HideObj (141)"
	case RecTypeSort:
		return "Sort (144)"
	case RecTypePalette:
		return "Palette (146)"
	case RecTypeSync:
		return "Sync (151)"
	case RecTypeLPr:
		return "LPr (152)"
	case RecTypeDxGCol:
		return "DxGCol (153)"
	case RecTypeFnGroupName:
		return "FnGroupName (154)"
	case RecTypeFilterMode:
		return "FilterMode (155)"
	case RecTypeBuiltInFnGroupCount:
		return "BuiltInFnGroupCount (156)"
	case RecTypeAutoFilterInfo:
		return "AutoFilterInfo (157)"
	case RecTypeAutoFilter:
		return "AutoFilter (158)"
	case RecTypeScl:
		return "Scl (160)"
	case RecTypeSetup:
		return "Setup (161)"
	case RecTypeScenMan:
		return "ScenMan (174)"
	case RecTypeSCENARIO:
		return "SCENARIO (175)"
	case RecTypeSxView:
		return "SxView (176)"
	case RecTypeSxvd:
		return "Sxvd (177)"
	case RecTypeSXVI:
		return "SXVI (178)"
	case RecTypeSxIvd:
		return "SxIvd (180)"
	case RecTypeSXLI:
		return "SXLI (181)"
	case RecTypeSXPI:
		return "SXPI (182)"
	case RecTypeDocRoute:
		return "DocRoute (184)"
	case RecTypeRecipName:
		return "RecipName (185)"
	case RecTypeMulRk:
		return "MulRk (189)"
	case RecTypeMulBlank:
		return "MulBlank (190)"
	case RecTypeMms:
		return "Mms (193)"
	case RecTypeSXDI:
		return "SXDI (197)"
	case RecTypeSXDB:
		return "SXDB (198)"
	case RecTypeSXFDB:
		return "SXFDB (199)"
	case RecTypeSXDBB:
		return "SXDBB (200)"
	case RecTypeSXNum:
		return "SXNum (201)"
	case RecTypeSxBool:
		return "SxBool (202)"
	case RecTypeSxErr:
		return "SxErr (203)"
	case RecTypeSXInt:
		return "SXInt (204)"
	case RecTypeSXString:
		return "SXString (205)"
	case RecTypeSXDtr:
		return "SXDtr (206)"
	case RecTypeSxNil:
		return "SxNil (207)"
	case RecTypeSXTbl:
		return "SXTbl (208)"
	case RecTypeSXTBRGIITM:
		return "SXTBRGIITM (209)"
	case RecTypeSxTbpg:
		return "SxTbpg (210)"
	case RecTypeObProj:
		return "ObProj (211)"
	case RecTypeSXStreamID:
		return "SXStreamID (213)"
	case RecTypeDBCell:
		return "DBCell (215)"
	case RecTypeSXRng:
		return "SXRng (216)"
	case RecTypeSxIsxoper:
		return "SxIsxoper (217)"
	case RecTypeBookBool:
		return "BookBool (218)"
	case RecTypeDbOrParamQry:
		return "DbOrParamQry (220)"
	case RecTypeScenarioProtect:
		return "ScenarioProtect (221)"
	case RecTypeOleObjectSize:
		return "OleObjectSize (222)"
	case RecTypeXF:
		return "XF (224)"
	case RecTypeInterfaceHdr:
		return "InterfaceHdr (225)"
	case RecTypeInterfaceEnd:
		return "InterfaceEnd (226)"
	case RecTypeSXVS:
		return "SXVS (227)"
	case RecTypeMergeCells:
		return "MergeCells (229)"
	case RecTypeBkHim:
		return "BkHim (233)"
	case RecTypeMsoDrawingGroup:
		return "MsoDrawingGroup (235)"
	case RecTypeMsoDrawing:
		return "MsoDrawing (236)"
	case RecTypeMsoDrawingSelection:
		return "MsoDrawingSelection (237)"
	case RecTypePhoneticInfo:
		return "PhoneticInfo (239)"
	case RecTypeSxRule:
		return "SxRule (240)"
	case RecTypeSXEx:
		return "SXEx (241)"
	case RecTypeSxFilt:
		return "SxFilt (242)"
	case RecTypeSxDXF:
		return "SxDXF (244)"
	case RecTypeSxItm:
		return "SxItm (245)"
	case RecTypeSxName:
		return "SxName (246)"
	case RecTypeSxSelect:
		return "SxSelect (247)"
	case RecTypeSXPair:
		return "SXPair (248)"
	case RecTypeSxFmla:
		return "SxFmla (249)"
	case RecTypeSxFormat:
		return "SxFormat (251)"
	case RecTypeSST:
		return "SST (252)"
	case RecTypeLabelSst:
		return "LabelSst (253)"
	case RecTypeExtSST:
		return "ExtSST (255)"
	case RecTypeSXVDEx:
		return "SXVDEx (256)"
	case RecTypeSXFormula:
		return "SXFormula (259)"
	case RecTypeSXDBEx:
		return "SXDBEx (290)"
	case RecTypeRRDInsDel:
		return "RRDInsDel (311)"
	case RecTypeRRDHead:
		return "RRDHead (312)"
	case RecTypeRRDChgCell:
		return "RRDChgCell (315)"
	case RecTypeRRTabID:
		return "RRTabID (317)"
	case RecTypeRRDRenSheet:
		return "RRDRenSheet (318)"
	case RecTypeRRSort:
		return "RRSort (319)"
	case RecTypeRRDMove:
		return "RRDMove (320)"
	case RecTypeRRFormat:
		return "RRFormat (330)"
	case RecTypeRRAutoFmt:
		return "RRAutoFmt (331)"
	case RecTypeRRInsertSh:
		return "RRInsertSh (333)"
	case RecTypeRRDMoveBegin:
		return "RRDMoveBegin (334)"
	case RecTypeRRDMoveEnd:
		return "RRDMoveEnd (335)"
	case RecTypeRRDInsDelBegin:
		return "RRDInsDelBegin (336)"
	case RecTypeRRDInsDelEnd:
		return "RRDInsDelEnd (337)"
	case RecTypeRRDConflict:
		return "RRDConflict (338)"
	case RecTypeRRDDefName:
		return "RRDDefName (339)"
	case RecTypeRRDRstEtxp:
		return "RRDRstEtxp (340)"
	case RecTypeLRng:
		return "LRng (351)"
	case RecTypeUsesELFs:
		return "UsesELFs (352)"
	case RecTypeDSF:
		return "DSF (353)"
	case RecTypeCUsr:
		return "CUsr (401)"
	case RecTypeCbUsr:
		return "CbUsr (402)"
	case RecTypeUsrInfo:
		return "UsrInfo (403)"
	case RecTypeUsrExcl:
		return "UsrExcl (404)"
	case RecTypeFileLock:
		return "FileLock (405)"
	case RecTypeRRDInfo:
		return "RRDInfo (406)"
	case RecTypeBCUsrs:
		return "BCUsrs (407)"
	case RecTypeUsrChk:
		return "UsrChk (408)"
	case RecTypeUserBView:
		return "UserBView (425)"
	case RecTypeUserSViewBegin:
		return "UserSViewBegin[Chart] (426)"
	case RecTypeUserSViewEnd:
		return "UserSViewEnd (427)"
	case RecTypeRRDUserView:
		return "RRDUserView (428)"
	case RecTypeQsi:
		return "Qsi (429)"
	case RecTypeSupBook:
		return "SupBook (430)"
	case RecTypeProt4Rev:
		return "Prot4Rev (431)"
	case RecTypeCondFmt:
		return "CondFmt (432)"
	case RecTypeCF:
		return "CF (433)"
	case RecTypeDVal:
		return "DVal (434)"
	case RecTypeDConBin:
		return "DConBin (437)"
	case RecTypeTxO:
		return "TxO (438)"
	case RecTypeRefreshAll:
		return "RefreshAll (439)"
	case RecTypeHLink:
		return "HLink (440)"
	case RecTypeLel:
		return "Lel (441)"
	case RecTypeCodeName:
		return "CodeName (442)"
	case RecTypeSXFDBType:
		return "SXFDBType (443)"
	case RecTypeProt4RevPass:
		return "Prot4RevPass (444)"
	case RecTypeObNoMacros:
		return "ObNoMacros (445)"
	case RecTypeDv:
		return "Dv (446)"
	case RecTypeExcel9File:
		return "Excel9File (448)"
	case RecTypeRecalcID:
		return "RecalcID (449)"
	case RecTypeEntExU2:
		return "EntExU2 (450)"
	case RecTypeDimensions:
		return "Dimensions (512)"
	case RecTypeBlank:
		return "Blank (513)"
	case RecTypeNumber:
		return "Number (515)"
	case RecTypeLabel:
		return "Label (516)"
	case RecTypeBoolErr:
		return "BoolErr (517)"
	case RecTypeString:
		return "String (519)"
	case RecTypeRow:
		return "Row (520)"
	case RecTypeIndex:
		return "Index (523)"
	case RecTypeArray:
		return "Array (545)"
	case RecTypeDefaultRowHeight:
		return "DefaultRowHeight (549)"
	case RecTypeTable:
		return "Table (566)"
	case RecTypeWindow2:
		return "Window2 (574)"
	case RecTypeRK:
		return "RK (638)"
	case RecTypeStyle:
		return "Style (659)"
	case RecTypeBigName:
		return "BigName (1048)"
	case RecTypeFormat:
		return "Format (1054)"
	case RecTypeContinueBigName:
		return "ContinueBigName (1084)"
	case RecTypeShrFmla:
		return "ShrFmla (1212)"
	case RecTypeHLinkTooltip:
		return "HLinkTooltip (2048)"
	case RecTypeWebPub:
		return "WebPub (2049)"
	case RecTypeQsiSXTag:
		return "QsiSXTag (2050)"
	case RecTypeDBQueryExt:
		return "DBQueryExt (2051)"
	case RecTypeExtString:
		return "ExtString (2052)"
	case RecTypeTxtQry:
		return "TxtQry (2053)"
	case RecTypeQsir:
		return "Qsir (2054)"
	case RecTypeQsif:
		return "Qsif (2055)"
	case RecTypeRRDTQSIF:
		return "RRDTQSIF (2056)"
	case RecTypeBOF:
		return "BOF (2057)"
	case RecTypeOleDbConn:
		return "OleDbConn (2058)"
	case RecTypeWOpt:
		return "WOpt (2059)"
	case RecTypeSXViewEx:
		return "SXViewEx (2060)"
	case RecTypeSXTH:
		return "SXTH (2061)"
	case RecTypeSXPIEx:
		return "SXPIEx (2062)"
	case RecTypeSXVDTEx:
		return "SXVDTEx (2063)"
	case RecTypeSXViewEx9:
		return "SXViewEx9 (2064)"
	case RecTypeContinueFrt:
		return "ContinueFrt (2066)"
	case RecTypeRealTimeData:
		return "RealTimeData (2067)"
	case RecTypeChartFrtInfo:
		return "ChartFrtInfo (2128)"
	case RecTypeFrtWrapper:
		return "FrtWrapper (2129)"
	case RecTypeStartBlock:
		return "StartBlock (2130)"
	case RecTypeEndBlock:
		return "EndBlock (2131)"
	case RecTypeStartObject:
		return "StartObject (2132)"
	case RecTypeEndObject:
		return "EndObject (2133)"
	case RecTypeCatLab:
		return "CatLab (2134)"
	case RecTypeYMult:
		return "YMult (2135)"
	case RecTypeSXViewLink:
		return "SXViewLink (2136)"
	case RecTypePivotChartBits:
		return "PivotChartBits (2137)"
	case RecTypeFrtFontList:
		return "FrtFontList (2138)"
	case RecTypeSheetExt:
		return "SheetExt (2146)"
	case RecTypeBookExt:
		return "BookExt (2147)"
	case RecTypeSXAddl:
		return "SXAddl (2148)"
	case RecTypeCrErr:
		return "CrErr (2149)"
	case RecTypeHFPicture:
		return "HFPicture (2150)"
	case RecTypeFeatHdr:
		return "FeatHdr (2151)"
	case RecTypeFeat:
		return "Feat (2152)"
	case RecTypeDataLabExt:
		return "DataLabExt (2154)"
	case RecTypeDataLabExtContents:
		return "DataLabExtContents (2155)"
	case RecTypeCellWatch:
		return "CellWatch (2156)"
	case RecTypeFeatHdr11:
		return "FeatHdr11 (2161)"
	case RecTypeFeature11:
		return "Feature11 (2162)"
	case RecTypeDropDownObjIds:
		return "DropDownObjIds (2164)"
	case RecTypeContinueFrt11:
		return "ContinueFrt11 (2165)"
	case RecTypeDConn:
		return "DConn (2166)"
	case RecTypeList12:
		return "List12 (2167)"
	case RecTypeFeature12:
		return "Feature12 (2168)"
	case RecTypeCondFmt12:
		return "CondFmt12 (2169)"
	case RecTypeCF12:
		return "CF12 (2170)"
	case RecTypeCFEx:
		return "CFEx (2171)"
	case RecTypeXFCRC:
		return "XFCRC (2172)"
	case RecTypeXFExt:
		return "XFExt (2173)"
	case RecTypeAutoFilter12:
		return "AutoFilter12 (2174)"
	case RecTypeContinueFrt12:
		return "ContinueFrt12 (2175)"
	case RecTypeMDTInfo:
		return "MDTInfo (2180)"
	case RecTypeMDXStr:
		return "MDXStr (2181)"
	case RecTypeMDXTuple:
		return "MDXTuple (2182)"
	case RecTypeMDXSet:
		return "MDXSet (2183)"
	case RecTypeMDXProp:
		return "MDXProp (2184)"
	case RecTypeMDXKPI:
		return "MDXKPI (2185)"
	case RecTypeMDB:
		return "MDB (2186)"
	case RecTypePLV:
		return "PLV (2187)"
	case RecTypeCompat12:
		return "Compat12 (2188)"
	case RecTypeDXF:
		return "DXF (2189)"
	case RecTypeTableStyles:
		return "TableStyles (2190)"
	case RecTypeTableStyle:
		return "TableStyle (2191)"
	case RecTypeTableStyleElement:
		return "TableStyleElement (2192)"
	case RecTypeStyleExt:
		return "StyleExt (2194)"
	case RecTypeNamePublish:
		return "NamePublish (2195)"
	case RecTypeNameCmt:
		return "NameCmt (2196)"
	case RecTypeSortData:
		return "SortData (2197)"
	case RecTypeTheme:
		return "Theme (2198)"
	case RecTypeGUIDTypeLib:
		return "GUIDTypeLib (2199)"
	case RecTypeFnGrp12:
		return "FnGrp12 (2200)"
	case RecTypeNameFnGrp12:
		return "NameFnGrp12 (2201)"
	case RecTypeMTRSettings:
		return "MTRSettings (2202)"
	case RecTypeCompressPictures:
		return "CompressPictures (2203)"
	case RecTypeHeaderFooter:
		return "HeaderFooter (2204)"
	case RecTypeCrtLayout12:
		return "CrtLayout12 (2205)"
	case RecTypeCrtMlFrt:
		return "CrtMlFrt (2206)"
	case RecTypeCrtMlFrtContinue:
		return "CrtMlFrtContinue (2207)"
	case RecTypeForceFullCalculation:
		return "ForceFullCalculation (2211)"
	case RecTypeShapePropsStream:
		return "ShapePropsStream (2212)"
	case RecTypeTextPropsStream:
		return "TextPropsStream (2213)"
	case RecTypeRichTextStream:
		return "RichTextStream (2214)"
	case RecTypeCrtLayout12A:
		return "CrtLayout12A (2215)"
	case RecTypeUnits:
		return "Units (4097)"
	case RecTypeChart:
		return "Chart (4098)"
	case RecTypeSeries:
		return "Series (4099)"
	case RecTypeDataFormat:
		return "DataFormat (4102)"
	case RecTypeLineFormat:
		return "LineFormat (4103)"
	case RecTypeMarkerFormat:
		return "MarkerFormat (4105)"
	case RecTypeAreaFormat:
		return "AreaFormat (4106)"
	case RecTypePieFormat:
		return "PieFormat (4107)"
	case RecTypeAttachedLabel:
		return "AttachedLabel (4108)"
	case RecTypeSeriesText:
		return "SeriesText (4109)"
	case RecTypeChartFormat:
		return "ChartFormat (4116)"
	case RecTypeLegend:
		return "Legend (4117)"
	case RecTypeSeriesList:
		return "SeriesList (4118)"
	case RecTypeBar:
		return "Bar (4119)"
	case RecTypeLine:
		return "Line (4120)"
	case RecTypePie:
		return "Pie (4121)"
	case RecTypeArea:
		return "Area (4122)"
	case RecTypeScatter:
		return "Scatter (4123)"
	case RecTypeCrtLine:
		return "CrtLine (4124)"
	case RecTypeAxis:
		return "Axis (4125)"
	case RecTypeTick:
		return "Tick (4126)"
	case RecTypeValueRange:
		return "ValueRange (4127)"
	case RecTypeCatSerRange:
		return "CatSerRange (4128)"
	case RecTypeAxisLine:
		return "AxisLine (4129)"
	case RecTypeCrtLink:
		return "CrtLink (4130)"
	case RecTypeDefaultText:
		return "DefaultText (4132)"
	case RecTypeText:
		return "Text (4133)"
	case RecTypeFontX:
		return "FontX (4134)"
	case RecTypeObjectLink:
		return "ObjectLink (4135)"
	case RecTypeFrame:
		return "Frame (4146)"
	case RecTypeBegin:
		return "Begin (4147)"
	case RecTypeEnd:
		return "End (4148)"
	case RecTypePlotArea:
		return "PlotArea (4149)"
	case RecTypeChart3d:
		return "Chart3d (4154)"
	case RecTypePicF:
		return "PicF (4156)"
	case RecTypeDropBar:
		return "DropBar (4157)"
	case RecTypeRadar:
		return "Radar (4158)"
	case RecTypeSurf:
		return "Surf (4159)"
	case RecTypeRadarArea:
		return "RadarArea (4160)"
	case RecTypeAxisParent:
		return "AxisParent (4161)"
	case RecTypeLegendException:
		return "LegendException (4163)"
	case RecTypeShtProps:
		return "ShtProps (4164)"
	case RecTypeSerToCrt:
		return "SerToCrt (4165)"
	case RecTypeAxesUsed:
		return "AxesUsed (4166)"
	case RecTypeSBaseRef:
		return "SBaseRef (4168)"
	case RecTypeSerParent:
		return "SerParent (4170)"
	case RecTypeSerAuxTrend:
		return "SerAuxTrend (4171)"
	case RecTypeIFmtRecord:
		return "IFmtRecord (4174)"
	case RecTypePos:
		return "Pos (4175)"
	case RecTypeAlRuns:
		return "AlRuns (4176)"
	case RecTypeBRAI:
		return "BRAI (4177)"
	case RecTypeSerAuxErrBar:
		return "SerAuxErrBar (4187)"
	case RecTypeClrtClient:
		return "ClrtClient (4188)"
	case RecTypeSerFmt:
		return "SerFmt (4189)"
	case RecTypeChart3DBarShape:
		return "Chart3DBarShape (4191)"
	case RecTypeFbi:
		return "Fbi (4192)"
	case RecTypeBopPop:
		return "BopPop (4193)"
	case RecTypeAxcExt:
		return "AxcExt (4194)"
	case RecTypeDat:
		return "Dat (4195)"
	case RecTypePlotGrowth:
		return "PlotGrowth (4196)"
	case RecTypeSIIndex:
		return "SIIndex (4197)"
	case RecTypeGelFrame:
		return "GelFrame (4198)"
	case RecTypeBopPopCustom:
		return "BopPopCustom (4199)"
	case RecTypeFbi2:
		return "Fbi2 (4200)"
	}
	return fmt.Sprintf("unknown (%d 0x%x)", uint16(r), uint16(r))
}
