using System;
using System.Collections.Generic;
using System.Xml;
using unvell.ReoGrid.Chart;
using unvell.ReoGrid.Drawing;
using unvell.ReoGrid.IO.OpenXML;
using unvell.ReoGrid.IO.OpenXML.Schema;

namespace unvell.ReoGrid.IO.Additional.Excel.FloatingObjects
{
    class Pie2DChartExporter : DrawingObjectExporterBase
    {
        #region DrawingObjectExporterBase

        public override void Export(Document doc, OpenXML.Schema.Worksheet sheet, OpenXML.Schema.Drawing drawing, Worksheet rgSheet, IDrawingObject exportObject)
        {
            if (CanExport(exportObject))
            {
                // Здесь будет сохранение pie chart
                WriteChart(doc, sheet, drawing, rgSheet, exportObject as Pie2DChart);
            }
            else
            {
                throw new ArgumentException("", nameof(exportObject));
            }
        }

        public override bool CanExport(IDrawingObject exportObject)
        {
            return exportObject is Pie2DChart;
        }

        #endregion

        private static void WriteChart(
            Document doc,
            OpenXML.Schema.Worksheet sheet,
            OpenXML.Schema.Drawing drawing,
            Worksheet rgSheet,
            Pie2DChart chart)
        {
            if (drawing.twoCellAnchors == null)
            {
                drawing.twoCellAnchors = new List<CT_TwoCellAnchor>();
            }

            string typeName = /*chart.GetFriendlyTypeName()*/ "Pie2DChart";

            drawing._typeObjectCount.TryGetValue(typeName, out var typeObjCount);
            typeObjCount++;

            drawing._typeObjectCount[typeName] = typeObjCount;

            // Сначала создаем chart для того чтобы получить rId на него
            var chartSpaceCreationResult = doc.CreateMediaChartSpace(sheet, drawing);

            var twoCellAnchor = new CT_TwoCellAnchor
            {
                from = CreateCellAnchorByLocation_FormicrosoftXsd(rgSheet, chart.Location),
                to = CreateCellAnchorByLocation_FormicrosoftXsd(rgSheet, new Graphics.Point(chart.Right, chart.Bottom)),
                Item = new CT_GraphicalObjectFrame
                {
                    macro = string.Empty,
                    nvGraphicFramePr = new CT_GraphicalObjectFrameNonVisual
                    {
                        cNvPr = new CT_NonVisualDrawingProps
                        {
                            id = (uint)drawing._drawingObjectCount++,           // Вот с этим параметром могут быть проблемы уточнить как он появляяется
                                                                                // Описан в 19.3.1.12
                                                                                // Этот идентификатор должен быть уникальным по всему документу
                            name = typeName + " " + typeObjCount,
                        },

                        cNvGraphicFramePr = new CT_NonVisualGraphicFrameProperties
                        {
                            graphicFrameLocks = new CT_GraphicalObjectFrameLocking()
                        },
                    },
                    xfrm = new CT_Transform2D
                    {
                        off = new CT_Point2D { x = 0, y = 0},
                        ext = new CT_PositiveSize2D { cx = 0, cy = 0},
                    },

                    graphic = new CT_GraphicalObject
                    {
                        graphicData = new CT_GraphicalObjectData
                        {
                            uri = OpenXMLNamespaces.Chart________,
                            chart = new CT_RelId
                            {
                                id = chartSpaceCreationResult.RId
                            }
                        }
                    },
                },
                clientData = new CT_AnchorClientData(),
            };

            // Далее надо заполнить chartSpaceCreationResult и куда то его 
            // записать так чтобы он сериализовался (изучить reogrid)
            drawing._chartSpace = chartSpaceCreationResult.Result;
            drawing.twoCellAnchors.Add(twoCellAnchor);

            FillChartSpace(chartSpaceCreationResult.Result, chart);
        }

        private static void FillChartSpace(CT_ChartSpace space, Pie2DChart chart)
        {
            space.date1904 = new CT_Boolean { val = false };
            space.lang = new CT_TextLanguageID
            {
                val = "ru-RU",
            };
            space.roundedCorners = new CT_Boolean { val = false };
            //пропущен <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
            space.chart = new CT_Chart
            {
                title = new CT_Title
                {
                    tx = new CT_Tx
                    {
                        Item = new CT_TextBody
                        {
                            bodyPr = new CT_TextBodyProperties
                            {
                                rot = 0,
                                rotSpecified = true,
                                spcFirstLastPara = true,
                                spcFirstLastParaSpecified = true,
                                vertOverflow = ST_TextVertOverflowType.ellipsis,
                                vertOverflowSpecified = true,
                                vert = ST_TextVerticalType.horz,
                                vertSpecified = true,
                                wrap = ST_TextWrappingType.square,
                                wrapSpecified = true,
                                anchor = ST_TextAnchoringType.ctr,
                                anchorSpecified = true,
                                anchorCtr = true,
                                anchorCtrSpecified = true
                            },
                            lstStyle = new CT_TextListStyle { },
                            p = new[]
                            {
                                new CT_TextParagraph
                                {
                                    pPr = new CT_TextParagraphProperties
                                    {
                                        defRPr = new CT_TextCharacterProperties
                                        {
                                            sz = 1400, szSpecified = true,
                                            b=false, bSpecified = true,
                                            i = false, iSpecified =  true,
                                            u = ST_TextUnderlineType.none, uSpecified = true,
                                            strike = ST_TextStrikeType.noStrike, strikeSpecified = true,
                                            kern = 1200, kernSpecified = true,
                                            spc = 0, spcSpecified = true,
                                            baseline = 0, baselineSpecified = true,
                                            solidFill = new CT_SolidColorFillProperties
                                            {
                                                schemeClr = new CT_SchemeColor
                                                {
                                                    val = ST_SchemeColorVal.tx1,
                                                    lumMod = new []{ new CT_Percentage{ val = 65000 }},
                                                    lumOff = new []{ new CT_Percentage{ val = 35000 }},
                                                }
                                            },
                                            latin = new CT_TextFont{typeface = @"+mn-lt"},
                                            ea = new CT_TextFont{typeface = @"+mn-ea"},
                                            cs = new CT_TextFont{typeface = @"+mn-cs"},
                                        },
                                    },
                                    r = new []
                                    {
                                        new CT_RegularTextRun
                                        {
                                            rPr = new CT_TextCharacterProperties{lang = "en-US"},
                                            t = chart.Title,
                                        }
                                    },
                                    endParaRPr = new CT_TextCharacterProperties{lang = "ru-RU"},
                                },
                            }
                        },
                    },
                    layout = new CT_Layout(),
                    overlay = new CT_Boolean { val = false },
                    spPr = new CT_ShapeProperties
                    {
                        noFill = new CT_NoFillProperties(),
                        ln = new CT_LineProperties { noFill = new CT_NoFillProperties() },
                        effectLst = new CT_EffectList()
                    },
                    txPr = new CT_TextBody
                    {
                        bodyPr = new CT_TextBodyProperties
                        {
                            rot = 0, rotSpecified = true,
                            spcFirstLastPara = true, spcFirstLastParaSpecified = true,
                            vertOverflow = ST_TextVertOverflowType.ellipsis, vertOverflowSpecified = true,
                            vert = ST_TextVerticalType.horz, vertSpecified = true,
                            wrap = ST_TextWrappingType.square, wrapSpecified = true,
                            anchor = ST_TextAnchoringType.ctr, anchorSpecified = true,
                            anchorCtr = true, anchorCtrSpecified = true,
                        },
                        lstStyle = new CT_TextListStyle(),
                        p = new[]
                        {
                            new CT_TextParagraph
                            {
                                pPr = new CT_TextParagraphProperties
                                {
                                    defRPr = new CT_TextCharacterProperties
                                    {
                                        sz = 1400, szSpecified = true,
                                        b=false, bSpecified = true,
                                        i = false, iSpecified =  true,
                                        u = ST_TextUnderlineType.none, uSpecified = true,
                                        strike = ST_TextStrikeType.noStrike, strikeSpecified = true,
                                        kern = 1200, kernSpecified = true,
                                        spc = 0, spcSpecified = true,
                                        baseline = 0, baselineSpecified = true,
                                        solidFill = new CT_SolidColorFillProperties
                                        {
                                            schemeClr = new CT_SchemeColor
                                            {
                                                val = ST_SchemeColorVal.tx1,
                                                lumMod = new []{ new CT_Percentage{ val = 65000 }},
                                                lumOff = new []{ new CT_Percentage{ val = 35000 }},
                                            }
                                        },
                                        latin = new CT_TextFont{typeface = @"+mn-lt"},
                                        ea = new CT_TextFont{typeface = @"+mn-ea"},
                                        cs = new CT_TextFont{typeface = @"+mn-cs"},
                                    },
                                },
                                endParaRPr = new CT_TextCharacterProperties{lang = "ru-RU"},
                            },
                        },
                    },
                },
                autoTitleDeleted = new CT_Boolean { val = false },
                plotArea = new CT_PlotArea
                {
                    layout = new CT_Layout(),
                    Items = new object[]
                    {
                        new CT_PieChart
                        {
                            varyColors = new CT_Boolean{val = true,},
                            ser = new CT_PieSer[]
                            {
                                CreatePieSer(space, chart)
                            },
                            dLbls = new CT_DLbls
                            {
                                ItemsElementName = new[]
                                {
                                    ItemsChoiceType2.showLegendKey,
                                    ItemsChoiceType2.showVal,
                                    ItemsChoiceType2.showCatName,
                                    ItemsChoiceType2.showSerName,
                                    ItemsChoiceType2.showPercent,
                                    ItemsChoiceType2.showBubbleSize,
                                    ItemsChoiceType2.showLeaderLines,
                                },
                                Items = new object[]
                                {
                                    new CT_Boolean{val = false},
                                    new CT_Boolean{val = false},//<c:showVal val="0"/>
                                    new CT_Boolean{val = false},// <c:showCatName val="0"/>
                                    new CT_Boolean{val = false},// <c:showSerName val="0"/>
                                    new CT_Boolean{val = false},// <c:showPercent val="0"/>
                                    new CT_Boolean{val = false},// <c:showBubbleSize val="0"/>
                                    new CT_Boolean{val = true},// <c:showLeaderLines val="1"/>
                                },
                            },
                            firstSliceAng = new CT_FirstSliceAng{val = 0},
                            // продолжить копировать D:\tasks\6459\excel_example\example\xl\charts\chart1.xml

                        },//CT_PieChart
                    },//Items
                    spPr = new CT_ShapeProperties
                    {
                        noFill = new CT_NoFillProperties { },
                        ln = new CT_LineProperties
                        {
                            noFill = new CT_NoFillProperties { }
                        },
                        effectLst = new CT_EffectList { }
                    },
                },
                legend = new CT_Legend
                {
                    legendPos = new CT_LegendPos { val = ST_LegendPos.b},
                    layout = new CT_Layout { },
                    overlay = new CT_Boolean { val = false},
                    spPr = new CT_ShapeProperties
                    {
                        noFill = new CT_NoFillProperties { },
                        ln = new CT_LineProperties
                        {
                            noFill = new CT_NoFillProperties { }
                        },
                        effectLst = new CT_EffectList { }
                    },
                    txPr = new CT_TextBody
                    {
                        bodyPr = new CT_TextBodyProperties
                        {
                            rot = 0, rotSpecified = true,
                            spcFirstLastPara = true, spcFirstLastParaSpecified = true,
                            vertOverflow = ST_TextVertOverflowType.ellipsis, vertOverflowSpecified = true,
                            vert = ST_TextVerticalType.horz, vertSpecified = true,
                            wrap = ST_TextWrappingType.square, wrapSpecified = true,
                            anchor = ST_TextAnchoringType.ctr, anchorSpecified = true,
                            anchorCtr = true, anchorCtrSpecified = true,
                        },
                        lstStyle = new CT_TextListStyle{},
                        p = new CT_TextParagraph[]
                        {
                            new CT_TextParagraph
                            {
                                pPr = new CT_TextParagraphProperties
                                {
                                    rtl = false, rtlSpecified = true,
                                    defRPr = new CT_TextCharacterProperties
                                    {
                                        sz = 900, szSpecified = true,
                                        b = false, bSpecified = true,
                                        i = false, iSpecified = true,
                                        u = ST_TextUnderlineType.none, uSpecified = true,
                                        strike = ST_TextStrikeType.noStrike, strikeSpecified = true,
                                        kern = 1200, kernSpecified = true,
                                        baseline = 0, baselineSpecified = true,
                                        solidFill = new CT_SolidColorFillProperties
                                        {
                                            schemeClr = new CT_SchemeColor
                                            {
                                                val =ST_SchemeColorVal.tx1,
                                                lumMod = new []{new CT_Percentage { val = 65000}},  // TODO тут массив а в файле значение
                                                lumOff = new []{new CT_Percentage { val = 35000}},  // TODO тут массив а в файле значение
                                            }
                                        },
                                        latin = new CT_TextFont{typeface = @"+mn-lt"},              // TODO непонятная константа
                                        ea = new CT_TextFont{typeface = @"+mn-ea"},              // TODO непонятная константа
                                        cs = new CT_TextFont{typeface = @"+mn-cs"},              // TODO непонятная константа
                                    },
                                },
                                endParaRPr = new CT_TextCharacterProperties{lang = @"ru-RU"} // TODO непонятно от чего зависит
                            }
                        }
                        
                    }
                },
                plotVisOnly = new CT_Boolean { val = true},
                dispBlanksAs = new CT_DispBlanksAs { val = ST_DispBlanksAs.gap},
                showDLblsOverMax = new CT_Boolean { val = false}
            };

            space.spPr = new CT_ShapeProperties
            {
                solidFill = new CT_SolidColorFillProperties
                {
                    schemeClr = new CT_SchemeColor { val = ST_SchemeColorVal.bg1}
                },
                ln = new CT_LineProperties
                {
                    w = 9525, wSpecified = true,
                    cap = ST_LineCap.flat, capSpecified = true,
                    cmpd = ST_CompoundLine.sng, cmpdSpecified = true,
                    algn= ST_PenAlignment.ctr, algnSpecified = true,
                    solidFill = new CT_SolidColorFillProperties
                    {
                        schemeClr = new CT_SchemeColor
                        {
                            val = ST_SchemeColorVal.tx1,
                            lumMod = new[] { new CT_Percentage { val = 15000 } },
                            lumOff = new[] { new CT_Percentage { val = 85000 } },
                        },
                    },
                    round = new CT_LineJoinRound {},
                },
                effectLst = new CT_EffectList()
            };
            space.txPr = new CT_TextBody
            {
                bodyPr = new CT_TextBodyProperties(),
                lstStyle = new CT_TextListStyle(),
                p = new CT_TextParagraph[]
                {
                    new CT_TextParagraph
                    {
                        pPr = new CT_TextParagraphProperties {defRPr = new CT_TextCharacterProperties()},
                        endParaRPr = new CT_TextCharacterProperties {lang = "ru-RU"}
                    },
                },
            };
            space.printSettings = new CT_PrintSettings
            {
                headerFooter = new CT_HeaderFooter(),
                pageMargins = new CT_PageMargins
                {
                    b = 0.75,
                    l = 0.7,
                    r = 0.7,
                    t = 0.75,
                    header = 0.3,
                    footer = 0.3
                },
            };
        }

        private static CT_PieSer CreatePieSer(CT_ChartSpace space, Pie2DChart chart)
        {
            // chart.DataSource.Worksheet
            string catFormula = "'Chart &amp; Drawing'!$B$2:$F$2"; // полный адрес (формула) к подписям
            string valFormula = "'Chart &amp; Drawing'!$B$6:$F$6"; // полный адрес (формула) к значениям

            return new CT_PieSer
            {
                idx = new CT_UnsignedInt {val = 0}, // порядковый номер
                order = new CT_UnsignedInt {val = 0}, // ? что означает, может порядок в котором следуют значения?
                dPt = CreateDataPoints(space, chart),
                cat = new CT_AxDataSource {Item = CreateNumRef(catFormula)},
                val = new CT_NumDataSource {Item = CreateNumRef(valFormula)},
            };
        }

        private static CT_NumRef CreateNumRef(string formula)
        {
            return new CT_NumRef
            {
                f = formula,
                numCache = GetNumData(formula)
            };
        }

        // TODO при переработке в качестве параметра использовать бизнес-объекты с адресом, а не формулу в строковом виде
        private static CT_NumData GetNumData(string formula)
        {
            uint count = 5;                          // TODO расчитывать по реальной формуле
            var values = new List<CT_NumVal>();
            for (uint i = 0; i < count; i++)
            {
                values.Add(GetNumVal(formula, i));
            }
            return new CT_NumData
            {
                formatCode = @"General", // ? просто константа
                ptCount = new CT_UnsignedInt { val = (uint)values.Count },
                pt = values.ToArray(),
            };
        }

        private static CT_NumVal GetNumVal(string formula, uint index)
        {
            return new CT_NumVal
            {
                idx = index,
                v = "2019", // TODO брать реальное значение из рабочего листа
            };
        }

        // TODO Расчет проводить по реальному графику
        private static CT_DPt[] CreateDataPoints(CT_ChartSpace space, Pie2DChart chart)                                 
        {
            
            return new[]
            {
                CreateDataPoint(space, chart, 0),
                CreateDataPoint(space, chart, 1),
                CreateDataPoint(space, chart, 2),
                CreateDataPoint(space, chart, 3),
                CreateDataPoint(space, chart, 4),
            };
        }

        private static CT_DPt CreateDataPoint(CT_ChartSpace space, Pie2DChart chart, uint index)
        {
            return new CT_DPt
            {
                idx = new CT_UnsignedInt { val = index },
                bubble3D = new CT_Boolean { val = false },
                spPr = new CT_ShapeProperties
                {
                    solidFill = new CT_SolidColorFillProperties
                    {
                        schemeClr = new CT_SchemeColor { val = GetCyclicSchemeColor(index) }, /* А что если цвета кончатся? */
                    },
                    ln = new CT_LineProperties
                    {
                        w = 19050,
                        wSpecified = true, // в каких единицах измерения?
                        solidFill = new CT_SolidColorFillProperties
                        {
                            schemeClr = new CT_SchemeColor { val = ST_SchemeColorVal.lt1 } // вроде не меняется в зависимости от series
                        }
                    }
                }
            };
        }

        private static ST_SchemeColorVal GetCyclicSchemeColor(uint index)
        {
            switch (index % 6)
            {
                case 0: return ST_SchemeColorVal.accent1;
                case 1: return ST_SchemeColorVal.accent2;
                case 2: return ST_SchemeColorVal.accent3;
                case 3: return ST_SchemeColorVal.accent4;
                case 4: return ST_SchemeColorVal.accent5;
                case 5: return ST_SchemeColorVal.accent6;
            }
            return ST_SchemeColorVal.accent1; // иначе ошибка "not all code paths return a value"
        }
    }
}
