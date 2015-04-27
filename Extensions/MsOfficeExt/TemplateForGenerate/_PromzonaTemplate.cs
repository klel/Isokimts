using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Ds = DocumentFormat.OpenXml.CustomXmlDataProperties;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using kimts;

namespace Extensions.MsOfficeExt.TemplateForGenerate
{
    public class _PromzonaTemplate
    {
        private OutboxDocTamplate Template; 
        // Creates a WordprocessingDocument.
        public void CreatePackage(string filePath, OutboxDocTamplate tmpl)
        {
            Template = tmpl;

            using (WordprocessingDocument package = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(WordprocessingDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1);

            ImagePart imagePart1 = mainDocumentPart1.AddNewPart<ImagePart>("image/png", "rId8");
            GenerateImagePart1Content(imagePart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId3");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            EndnotesPart endnotesPart1 = mainDocumentPart1.AddNewPart<EndnotesPart>("rId7");
            GenerateEndnotesPart1Content(endnotesPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId12");
            GenerateThemePart1Content(themePart1);

            NumberingDefinitionsPart numberingDefinitionsPart1 = mainDocumentPart1.AddNewPart<NumberingDefinitionsPart>("rId2");
            GenerateNumberingDefinitionsPart1Content(numberingDefinitionsPart1);

            CustomXmlPart customXmlPart1 = mainDocumentPart1.AddNewPart<CustomXmlPart>("application/xml", "rId1");
            GenerateCustomXmlPart1Content(customXmlPart1);

            CustomXmlPropertiesPart customXmlPropertiesPart1 = customXmlPart1.AddNewPart<CustomXmlPropertiesPart>("rId1");
            GenerateCustomXmlPropertiesPart1Content(customXmlPropertiesPart1);

            FootnotesPart footnotesPart1 = mainDocumentPart1.AddNewPart<FootnotesPart>("rId6");
            GenerateFootnotesPart1Content(footnotesPart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId11");
            GenerateFontTablePart1Content(fontTablePart1);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId5");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            FooterPart footerPart1 = mainDocumentPart1.AddNewPart<FooterPart>("rId10");
            GenerateFooterPart1Content(footerPart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId4");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            documentSettingsPart1.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate", new System.Uri("file:///C:\\TamplatePlantZone9Panel.dotx", System.UriKind.Absolute), "rId1");
            FooterPart footerPart2 = mainDocumentPart1.AddNewPart<FooterPart>("rId9");
            GenerateFooterPart2Content(footerPart2);

            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Template template1 = new Ap.Template();
            template1.Text = "TamplatePlantZone9Panel.dotx";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "9";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "1";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "103";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "590";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Word";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "4";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "1";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Название";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "ОАО «ГАЗПРОМ»";

            vTVector2.Append(vTLPSTR2);

            titlesOfParts1.Append(vTVector2);
            Ap.Company company1 = new Ap.Company();
            company1.Text = "ООО \"Ноябрьскгаздобыча\"";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "692";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "12.0000";

            properties1.Append(template1);
            properties1.Append(totalTime1);
            properties1.Append(pages1);
            properties1.Append(words1);
            properties1.Append(characters1);
            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(lines1);
            properties1.Append(paragraphs1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(charactersWithSpaces1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of mainDocumentPart1.
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1)
        {
            Document document1 = new Document();
            document1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            Body body1 = new Body();

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "9860", Type = TableWidthUnitValues.Dxa };
            TableJustification tableJustification1 = new TableJustification() { Val = TableRowAlignmentValues.Center };
            TableLayout tableLayout1 = new TableLayout() { Type = TableLayoutValues.Fixed };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 0, Type = TableWidthValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 0, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);
            TableLook tableLook1 = new TableLook() { Val = "01E0" };

            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableJustification1);
            tableProperties1.Append(tableLayout1);
            tableProperties1.Append(tableCellMarginDefault1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "538" };
            GridColumn gridColumn2 = new GridColumn() { Width = "1417" };
            GridColumn gridColumn3 = new GridColumn() { Width = "590" };
            GridColumn gridColumn4 = new GridColumn() { Width = "1697" };
            GridColumn gridColumn5 = new GridColumn() { Width = "972" };
            GridColumn gridColumn6 = new GridColumn() { Width = "4646" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);
            tableGrid1.Append(gridColumn4);
            tableGrid1.Append(gridColumn5);
            tableGrid1.Append(gridColumn6);

            TableRow tableRow1 = new TableRow() { RsidTableRowAddition = "008A46A1", RsidTableRowProperties = "004D6D95" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)720U };
            TableJustification tableJustification2 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            tableRowProperties1.Append(tableRowHeight1);
            tableRowProperties1.Append(tableJustification2);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "4242", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan1 = new GridSpan() { Val = 4 };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(gridSpan1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "00520064" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE1 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN1 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent1 = new AdjustRightIndent() { Val = false };
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = "480", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color1 = new Color() { Val = "000000" };
            FontSize fontSize1 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(color1);
            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

            paragraphProperties1.Append(autoSpaceDE1);
            paragraphProperties1.Append(autoSpaceDN1);
            paragraphProperties1.Append(adjustRightIndent1);
            paragraphProperties1.Append(spacingBetweenLines1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            NoProof noProof1 = new NoProof();
            Color color2 = new Color() { Val = "000000" };
            FontSize fontSize2 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "20" };

            runProperties1.Append(runFonts2);
            runProperties1.Append(noProof1);
            runProperties1.Append(color2);
            runProperties1.Append(fontSize2);
            runProperties1.Append(fontSizeComplexScript2);

            Drawing drawing1 = new Drawing();

            Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251657728U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true };
            Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
            Wp.PositionOffset positionOffset1 = new Wp.PositionOffset();
            positionOffset1.Text = "1233170";

            horizontalPosition1.Append(positionOffset1);

            Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset2 = new Wp.PositionOffset();
            positionOffset2.Text = "-209550";

            verticalPosition1.Append(positionOffset2);
            Wp.Extent extent1 = new Wp.Extent() { Cx = 394970L, Cy = 668020L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 5080L, BottomEdge = 0L };
            Wp.WrapNone wrapNone1 = new Wp.WrapNone();
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)3U, Name = "Рисунок 13", Description = "015_1-1-2" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture1 = new Pic.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Рисунок 13", Description = "015_1-1-2" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Pic.BlipFill blipFill1 = new Pic.BlipFill();
            A.Blip blip1 = new A.Blip() { Embed = "rId8", CompressionState = A.BlipCompressionValues.Print };
            A.SourceRectangle sourceRectangle1 = new A.SourceRectangle();

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(sourceRectangle1);
            blipFill1.Append(stretch1);

            Pic.ShapeProperties shapeProperties1 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 394970L, Cy = 668020L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);
            A.NoFill noFill1 = new A.NoFill();

            A.Outline outline1 = new A.Outline() { Width = 9525 };
            A.NoFill noFill2 = new A.NoFill();
            A.Miter miter1 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd1 = new A.HeadEnd();
            A.TailEnd tailEnd1 = new A.TailEnd();

            outline1.Append(noFill2);
            outline1.Append(miter1);
            outline1.Append(headEnd1);
            outline1.Append(tailEnd1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(noFill1);
            shapeProperties1.Append(outline1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapNone1);
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic1);

            drawing1.Append(anchor1);

            run1.Append(runProperties1);
            run1.Append(drawing1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "972", Type = TableWidthUnitValues.Dxa };

            tableCellProperties2.Append(tableCellWidth2);
            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "003516C1", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "4646", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Bottom };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(tableCellVerticalAlignment1);
            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "000C6B6F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004E6F6C", RsidRunAdditionDefault = "008A46A1" };

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);

            TableRow tableRow2 = new TableRow() { RsidTableRowAddition = "008A46A1", RsidTableRowProperties = "004D6D95" };

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            TableRowHeight tableRowHeight2 = new TableRowHeight() { Val = (UInt32Value)1683U };
            TableJustification tableJustification3 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            tableRowProperties2.Append(tableRowHeight2);
            tableRowProperties2.Append(tableJustification3);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "4242", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan2 = new GridSpan() { Val = 4 };

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "0099FF", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders1.Append(bottomBorder1);

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(gridSpan2);
            tableCellProperties4.Append(tableCellBorders1);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE2 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN2 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent2 = new AdjustRightIndent() { Val = false };
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color3 = new Color() { Val = "000000" };
            FontSize fontSize3 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties2.Append(runFonts3);
            paragraphMarkRunProperties2.Append(color3);
            paragraphMarkRunProperties2.Append(fontSize3);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript3);

            paragraphProperties2.Append(autoSpaceDE2);
            paragraphProperties2.Append(autoSpaceDN2);
            paragraphProperties2.Append(adjustRightIndent2);
            paragraphProperties2.Append(spacingBetweenLines2);
            paragraphProperties2.Append(justification1);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run2 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color4 = new Color() { Val = "231F20" };
            FontSize fontSize4 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "20" };

            runProperties2.Append(runFonts4);
            runProperties2.Append(color4);
            runProperties2.Append(fontSize4);
            runProperties2.Append(fontSizeComplexScript4);
            Text text1 = new Text();
            text1.Text = "ОАО";

            run2.Append(runProperties2);
            run2.Append(text1);

            Run run3 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color5 = new Color() { Val = "231F20" };
            FontSize fontSize5 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "20" };

            runProperties3.Append(runFonts5);
            runProperties3.Append(color5);
            runProperties3.Append(fontSize5);
            runProperties3.Append(fontSizeComplexScript5);
            Text text2 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text2.Text = " «";

            run3.Append(runProperties3);
            run3.Append(text2);

            Run run4 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color6 = new Color() { Val = "231F20" };
            FontSize fontSize6 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "20" };

            runProperties4.Append(runFonts6);
            runProperties4.Append(color6);
            runProperties4.Append(fontSize6);
            runProperties4.Append(fontSizeComplexScript6);
            Text text3 = new Text();
            text3.Text = "ГАЗПРОМ";

            run4.Append(runProperties4);
            run4.Append(text3);

            Run run5 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts7 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color7 = new Color() { Val = "231F20" };
            FontSize fontSize7 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "20" };

            runProperties5.Append(runFonts7);
            runProperties5.Append(color7);
            runProperties5.Append(fontSize7);
            runProperties5.Append(fontSizeComplexScript7);
            Text text4 = new Text();
            text4.Text = "»";

            run5.Append(runProperties5);
            run5.Append(text4);

            paragraph4.Append(paragraphProperties2);
            paragraph4.Append(run2);
            paragraph4.Append(run3);
            paragraph4.Append(run4);
            paragraph4.Append(run5);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "00BB50A2", RsidParagraphAddition = "00BB50A2", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE3 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN3 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent3 = new AdjustRightIndent() { Val = false };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts8 = new RunFonts() { ComplexScript = "HeliosCond-Bold", AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi };
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            Color color8 = new Color() { Val = "231F20" };
            FontSize fontSize8 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties3.Append(runFonts8);
            paragraphMarkRunProperties3.Append(bold1);
            paragraphMarkRunProperties3.Append(boldComplexScript1);
            paragraphMarkRunProperties3.Append(color8);
            paragraphMarkRunProperties3.Append(fontSize8);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript8);

            paragraphProperties3.Append(autoSpaceDE3);
            paragraphProperties3.Append(autoSpaceDN3);
            paragraphProperties3.Append(adjustRightIndent3);
            paragraphProperties3.Append(justification2);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run6 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties6 = new RunProperties();
            RunFonts runFonts9 = new RunFonts() { Ascii = "HeliosCond-Bold", HighAnsi = "HeliosCond-Bold", ComplexScript = "HeliosCond-Bold" };
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            Color color9 = new Color() { Val = "231F20" };
            FontSize fontSize9 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "28" };

            runProperties6.Append(runFonts9);
            runProperties6.Append(bold2);
            runProperties6.Append(boldComplexScript2);
            runProperties6.Append(color9);
            runProperties6.Append(fontSize9);
            runProperties6.Append(fontSizeComplexScript9);
            Text text5 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text5.Text = "ОБЩЕСТВО С ";

            run6.Append(runProperties6);
            run6.Append(text5);
            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run7 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties7 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "HeliosCond-Bold", HighAnsi = "HeliosCond-Bold", ComplexScript = "HeliosCond-Bold" };
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            Color color10 = new Color() { Val = "231F20" };
            FontSize fontSize10 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "28" };

            runProperties7.Append(runFonts10);
            runProperties7.Append(bold3);
            runProperties7.Append(boldComplexScript3);
            runProperties7.Append(color10);
            runProperties7.Append(fontSize10);
            runProperties7.Append(fontSizeComplexScript10);
            Text text6 = new Text();
            text6.Text = "ОГРАНИЧЕННОЙ";

            run7.Append(runProperties7);
            run7.Append(text6);
            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            paragraph5.Append(paragraphProperties3);
            paragraph5.Append(run6);
            paragraph5.Append(proofError1);
            paragraph5.Append(run7);
            paragraph5.Append(proofError2);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE4 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN4 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent4 = new AdjustRightIndent() { Val = false };
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts11 = new RunFonts() { Ascii = "HeliosCond-Bold", HighAnsi = "HeliosCond-Bold", ComplexScript = "HeliosCond-Bold" };
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            Color color11 = new Color() { Val = "231F20" };
            FontSize fontSize11 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties4.Append(runFonts11);
            paragraphMarkRunProperties4.Append(bold4);
            paragraphMarkRunProperties4.Append(boldComplexScript4);
            paragraphMarkRunProperties4.Append(color11);
            paragraphMarkRunProperties4.Append(fontSize11);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript11);

            paragraphProperties4.Append(autoSpaceDE4);
            paragraphProperties4.Append(autoSpaceDN4);
            paragraphProperties4.Append(adjustRightIndent4);
            paragraphProperties4.Append(justification3);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run8 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts12 = new RunFonts() { Ascii = "HeliosCond-Bold", HighAnsi = "HeliosCond-Bold", ComplexScript = "HeliosCond-Bold" };
            Bold bold5 = new Bold();
            BoldComplexScript boldComplexScript5 = new BoldComplexScript();
            Color color12 = new Color() { Val = "231F20" };
            FontSize fontSize12 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "28" };

            runProperties8.Append(runFonts12);
            runProperties8.Append(bold5);
            runProperties8.Append(boldComplexScript5);
            runProperties8.Append(color12);
            runProperties8.Append(fontSize12);
            runProperties8.Append(fontSizeComplexScript12);
            Text text7 = new Text();
            text7.Text = "ОТВЕТСТВЕННОСТЬЮ";

            run8.Append(runProperties8);
            run8.Append(text7);

            paragraph6.Append(paragraphProperties4);
            paragraph6.Append(run8);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "00D82AD6", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE5 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN5 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent5 = new AdjustRightIndent() { Val = false };
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts13 = new RunFonts() { ComplexScript = "HeliosCond-Bold" };
            Bold bold6 = new Bold();
            BoldComplexScript boldComplexScript6 = new BoldComplexScript();
            Color color13 = new Color() { Val = "231F20" };
            FontSize fontSize13 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties5.Append(runFonts13);
            paragraphMarkRunProperties5.Append(bold6);
            paragraphMarkRunProperties5.Append(boldComplexScript6);
            paragraphMarkRunProperties5.Append(color13);
            paragraphMarkRunProperties5.Append(fontSize13);
            paragraphMarkRunProperties5.Append(fontSizeComplexScript13);

            paragraphProperties5.Append(autoSpaceDE5);
            paragraphProperties5.Append(autoSpaceDN5);
            paragraphProperties5.Append(adjustRightIndent5);
            paragraphProperties5.Append(spacingBetweenLines3);
            paragraphProperties5.Append(justification4);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run9 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts14 = new RunFonts() { Ascii = "HeliosCond-Bold", HighAnsi = "HeliosCond-Bold", ComplexScript = "HeliosCond-Bold" };
            Bold bold7 = new Bold();
            BoldComplexScript boldComplexScript7 = new BoldComplexScript();
            Color color14 = new Color() { Val = "231F20" };
            FontSize fontSize14 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "28" };

            runProperties9.Append(runFonts14);
            runProperties9.Append(bold7);
            runProperties9.Append(boldComplexScript7);
            runProperties9.Append(color14);
            runProperties9.Append(fontSize14);
            runProperties9.Append(fontSizeComplexScript14);
            Text text8 = new Text();
            text8.Text = "«ГАЗПРОМ ДОБЫЧА НОЯБРЬСК»";

            run9.Append(runProperties9);
            run9.Append(text8);

            paragraph7.Append(paragraphProperties5);
            paragraph7.Append(run9);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE6 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN6 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent6 = new AdjustRightIndent() { Val = false };
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Justification justification5 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts15 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color15 = new Color() { Val = "231F20" };
            FontSize fontSize15 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties6.Append(runFonts15);
            paragraphMarkRunProperties6.Append(color15);
            paragraphMarkRunProperties6.Append(fontSize15);
            paragraphMarkRunProperties6.Append(fontSizeComplexScript15);

            paragraphProperties6.Append(autoSpaceDE6);
            paragraphProperties6.Append(autoSpaceDN6);
            paragraphProperties6.Append(adjustRightIndent6);
            paragraphProperties6.Append(spacingBetweenLines4);
            paragraphProperties6.Append(justification5);
            paragraphProperties6.Append(paragraphMarkRunProperties6);

            Run run10 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts16 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color16 = new Color() { Val = "231F20" };
            FontSize fontSize16 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "20" };

            runProperties10.Append(runFonts16);
            runProperties10.Append(color16);
            runProperties10.Append(fontSize16);
            runProperties10.Append(fontSizeComplexScript16);
            Text text9 = new Text();
            text9.Text = "(";

            run10.Append(runProperties10);
            run10.Append(text9);

            Run run11 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts17 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color17 = new Color() { Val = "231F20" };
            FontSize fontSize17 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "20" };

            runProperties11.Append(runFonts17);
            runProperties11.Append(color17);
            runProperties11.Append(fontSize17);
            runProperties11.Append(fontSizeComplexScript17);
            Text text10 = new Text();
            text10.Text = "ООО";

            run11.Append(runProperties11);
            run11.Append(text10);

            Run run12 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts18 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color18 = new Color() { Val = "231F20" };
            FontSize fontSize18 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "20" };

            runProperties12.Append(runFonts18);
            runProperties12.Append(color18);
            runProperties12.Append(fontSize18);
            runProperties12.Append(fontSizeComplexScript18);
            Text text11 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text11.Text = " «";

            run12.Append(runProperties12);
            run12.Append(text11);

            Run run13 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties13 = new RunProperties();
            RunFonts runFonts19 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color19 = new Color() { Val = "231F20" };
            FontSize fontSize19 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "20" };

            runProperties13.Append(runFonts19);
            runProperties13.Append(color19);
            runProperties13.Append(fontSize19);
            runProperties13.Append(fontSizeComplexScript19);
            Text text12 = new Text();
            text12.Text = "Газпром";

            run13.Append(runProperties13);
            run13.Append(text12);

            Run run14 = new Run() { RsidRunAddition = "00BB50A2" };

            RunProperties runProperties14 = new RunProperties();
            RunFonts runFonts20 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color20 = new Color() { Val = "231F20" };
            FontSize fontSize20 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "20" };
            Languages languages1 = new Languages() { Val = "en-US" };

            runProperties14.Append(runFonts20);
            runProperties14.Append(color20);
            runProperties14.Append(fontSize20);
            runProperties14.Append(fontSizeComplexScript20);
            runProperties14.Append(languages1);
            Text text13 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text13.Text = " ";

            run14.Append(runProperties14);
            run14.Append(text13);

            Run run15 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties15 = new RunProperties();
            RunFonts runFonts21 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color21 = new Color() { Val = "231F20" };
            FontSize fontSize21 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "20" };

            runProperties15.Append(runFonts21);
            runProperties15.Append(color21);
            runProperties15.Append(fontSize21);
            runProperties15.Append(fontSizeComplexScript21);
            Text text14 = new Text();
            text14.Text = "добыча";

            run15.Append(runProperties15);
            run15.Append(text14);

            Run run16 = new Run() { RsidRunAddition = "00BB50A2" };

            RunProperties runProperties16 = new RunProperties();
            RunFonts runFonts22 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color22 = new Color() { Val = "231F20" };
            FontSize fontSize22 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "20" };
            Languages languages2 = new Languages() { Val = "en-US" };

            runProperties16.Append(runFonts22);
            runProperties16.Append(color22);
            runProperties16.Append(fontSize22);
            runProperties16.Append(fontSizeComplexScript22);
            runProperties16.Append(languages2);
            Text text15 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text15.Text = " ";

            run16.Append(runProperties16);
            run16.Append(text15);

            Run run17 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties17 = new RunProperties();
            RunFonts runFonts23 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color23 = new Color() { Val = "231F20" };
            FontSize fontSize23 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "20" };

            runProperties17.Append(runFonts23);
            runProperties17.Append(color23);
            runProperties17.Append(fontSize23);
            runProperties17.Append(fontSizeComplexScript23);
            Text text16 = new Text();
            text16.Text = "Ноябрьск";

            run17.Append(runProperties17);
            run17.Append(text16);

            Run run18 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties18 = new RunProperties();
            RunFonts runFonts24 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color24 = new Color() { Val = "231F20" };
            FontSize fontSize24 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "20" };

            runProperties18.Append(runFonts24);
            runProperties18.Append(color24);
            runProperties18.Append(fontSize24);
            runProperties18.Append(fontSizeComplexScript24);
            Text text17 = new Text();
            text17.Text = "»)";

            run18.Append(runProperties18);
            run18.Append(text17);

            paragraph8.Append(paragraphProperties6);
            paragraph8.Append(run10);
            paragraph8.Append(run11);
            paragraph8.Append(run12);
            paragraph8.Append(run13);
            paragraph8.Append(run14);
            paragraph8.Append(run15);
            paragraph8.Append(run16);
            paragraph8.Append(run17);
            paragraph8.Append(run18);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE7 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN7 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent7 = new AdjustRightIndent() { Val = false };
            Justification justification6 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            RunFonts runFonts25 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Bold bold8 = new Bold();
            Color color25 = new Color() { Val = "231F20" };
            FontSize fontSize25 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties7.Append(runFonts25);
            paragraphMarkRunProperties7.Append(bold8);
            paragraphMarkRunProperties7.Append(color25);
            paragraphMarkRunProperties7.Append(fontSize25);
            paragraphMarkRunProperties7.Append(fontSizeComplexScript25);

            paragraphProperties7.Append(autoSpaceDE7);
            paragraphProperties7.Append(autoSpaceDN7);
            paragraphProperties7.Append(adjustRightIndent7);
            paragraphProperties7.Append(justification6);
            paragraphProperties7.Append(paragraphMarkRunProperties7);

            Run run19 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties19 = new RunProperties();
            RunFonts runFonts26 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Bold bold9 = new Bold();
            NoProof noProof2 = new NoProof();

            runProperties19.Append(runFonts26);
            runProperties19.Append(bold9);
            runProperties19.Append(noProof2);
            Text text18 = new Text();
            text18.Text = "УПРАВЛЕНИЕ ОРГАНИЗАЦИИ РЕМОНТА, РЕКОНСТРУКЦИИ И СТРОИТЕЛЬСТВА ОСНОВНЫХ ФОНДОВ";

            run19.Append(runProperties19);
            run19.Append(text18);

            paragraph9.Append(paragraphProperties7);
            paragraph9.Append(run19);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE8 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN8 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent8 = new AdjustRightIndent() { Val = false };
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Justification justification7 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            NoProof noProof3 = new NoProof();
            FontSize fontSize26 = new FontSize() { Val = "16" };

            paragraphMarkRunProperties8.Append(noProof3);
            paragraphMarkRunProperties8.Append(fontSize26);

            paragraphProperties8.Append(autoSpaceDE8);
            paragraphProperties8.Append(autoSpaceDN8);
            paragraphProperties8.Append(adjustRightIndent8);
            paragraphProperties8.Append(spacingBetweenLines5);
            paragraphProperties8.Append(justification7);
            paragraphProperties8.Append(paragraphMarkRunProperties8);

            paragraph10.Append(paragraphProperties8);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph4);
            tableCell4.Append(paragraph5);
            tableCell4.Append(paragraph6);
            tableCell4.Append(paragraph7);
            tableCell4.Append(paragraph8);
            tableCell4.Append(paragraph9);
            tableCell4.Append(paragraph10);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "972", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge1 = new VerticalMerge() { Val = MergedCellValues.Restart };

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(verticalMerge1);
            Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "003516C1", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph11);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "4646", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge2 = new VerticalMerge() { Val = MergedCellValues.Restart };

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(verticalMerge2);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            Justification justification8 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            Bold bold10 = new Bold();
            FontSize fontSize27 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties9.Append(bold10);
            paragraphMarkRunProperties9.Append(fontSize27);
            paragraphMarkRunProperties9.Append(fontSizeComplexScript26);

            paragraphProperties9.Append(justification8);
            paragraphProperties9.Append(paragraphMarkRunProperties9);

            paragraph12.Append(paragraphProperties9);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            Justification justification9 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            Bold bold11 = new Bold();
            FontSize fontSize28 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties10.Append(bold11);
            paragraphMarkRunProperties10.Append(fontSize28);
            paragraphMarkRunProperties10.Append(fontSizeComplexScript27);

            paragraphProperties10.Append(justification9);
            paragraphProperties10.Append(paragraphMarkRunProperties10);

            paragraph13.Append(paragraphProperties10);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "00C71E0B", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "00BB50A2" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "a9" };
            Justification justification10 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            Bold bold12 = new Bold();
            BoldComplexScript boldComplexScript8 = new BoldComplexScript();
            FontSize fontSize29 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "28" };
            Languages languages3 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties11.Append(bold12);
            paragraphMarkRunProperties11.Append(boldComplexScript8);
            paragraphMarkRunProperties11.Append(fontSize29);
            paragraphMarkRunProperties11.Append(fontSizeComplexScript28);
            paragraphMarkRunProperties11.Append(languages3);

            paragraphProperties11.Append(paragraphStyleId1);
            paragraphProperties11.Append(justification10);
            paragraphProperties11.Append(paragraphMarkRunProperties11);
            ProofError proofError3 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run20 = new Run();

            RunProperties runProperties20 = new RunProperties();
            Bold bold13 = new Bold();
            BoldComplexScript boldComplexScript9 = new BoldComplexScript();
            FontSize fontSize30 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "28" };
            Languages languages4 = new Languages() { Val = "en-US" };

            runProperties20.Append(bold13);
            runProperties20.Append(boldComplexScript9);
            runProperties20.Append(fontSize30);
            runProperties20.Append(fontSizeComplexScript29);
            runProperties20.Append(languages4);
            Text text19 = new Text();
            text19.Text = Template.RecieverPost;

            run20.Append(runProperties20);
            run20.Append(text19);
            ProofError proofError4 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph14.Append(paragraphProperties11);
            paragraph14.Append(proofError3);
            paragraph14.Append(run20);
            paragraph14.Append(proofError4);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00BC2F1A", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "00C71E0B" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "a9" };
            Justification justification11 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            Bold bold14 = new Bold();
            FontSize fontSize31 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties12.Append(bold14);
            paragraphMarkRunProperties12.Append(fontSize31);
            paragraphMarkRunProperties12.Append(fontSizeComplexScript30);

            paragraphProperties12.Append(paragraphStyleId2);
            paragraphProperties12.Append(justification11);
            paragraphProperties12.Append(paragraphMarkRunProperties12);

            Run run21 = new Run();

            RunProperties runProperties21 = new RunProperties();
            Bold bold15 = new Bold();
            BoldComplexScript boldComplexScript10 = new BoldComplexScript();
            FontSize fontSize32 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "28" };
            Languages languages5 = new Languages() { Val = "en-US" };

            runProperties21.Append(bold15);
            runProperties21.Append(boldComplexScript10);
            runProperties21.Append(fontSize32);
            runProperties21.Append(fontSizeComplexScript31);
            runProperties21.Append(languages5);
            Text text20 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text20.Text = " ";

            run21.Append(runProperties21);
            run21.Append(text20);
            ProofError proofError5 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run22 = new Run() { RsidRunAddition = "00CC3955" };

            RunProperties runProperties22 = new RunProperties();
            Bold bold16 = new Bold();
            BoldComplexScript boldComplexScript11 = new BoldComplexScript();
            FontSize fontSize33 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "28" };
            Languages languages6 = new Languages() { Val = "en-US" };

            runProperties22.Append(bold16);
            runProperties22.Append(boldComplexScript11);
            runProperties22.Append(fontSize33);
            runProperties22.Append(fontSizeComplexScript32);
            runProperties22.Append(languages6);
            Text text21 = new Text();
            text21.Text = Template.RecieverOrg;

            run22.Append(runProperties22);
            run22.Append(text21);
            ProofError proofError6 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph15.Append(paragraphProperties12);
            paragraph15.Append(run21);
            paragraph15.Append(proofError5);
            paragraph15.Append(run22);
            paragraph15.Append(proofError6);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "00527607", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            Justification justification12 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            Bold bold17 = new Bold();
            FontSize fontSize34 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties13.Append(bold17);
            paragraphMarkRunProperties13.Append(fontSize34);
            paragraphMarkRunProperties13.Append(fontSizeComplexScript33);

            paragraphProperties13.Append(justification12);
            paragraphProperties13.Append(paragraphMarkRunProperties13);

            paragraph16.Append(paragraphProperties13);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "00BB50A2" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            Justification justification13 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            Bold bold18 = new Bold();
            FontSize fontSize35 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "28" };
            Languages languages7 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties14.Append(bold18);
            paragraphMarkRunProperties14.Append(fontSize35);
            paragraphMarkRunProperties14.Append(fontSizeComplexScript34);
            paragraphMarkRunProperties14.Append(languages7);

            paragraphProperties14.Append(justification13);
            paragraphProperties14.Append(paragraphMarkRunProperties14);
            ProofError proofError7 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run23 = new Run();

            RunProperties runProperties23 = new RunProperties();
            Bold bold19 = new Bold();
            FontSize fontSize36 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "28" };
            Languages languages8 = new Languages() { Val = "en-US" };

            runProperties23.Append(bold19);
            runProperties23.Append(fontSize36);
            runProperties23.Append(fontSizeComplexScript35);
            runProperties23.Append(languages8);
            Text text22 = new Text();
            text22.Text = Template.RecieverInitials;

            run23.Append(runProperties23);
            run23.Append(text22);
            ProofError proofError8 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph17.Append(paragraphProperties14);
            paragraph17.Append(proofError7);
            paragraph17.Append(run23);
            paragraph17.Append(proofError8);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();
            Justification justification14 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            FontSize fontSize37 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties15.Append(fontSize37);
            paragraphMarkRunProperties15.Append(fontSizeComplexScript36);

            paragraphProperties15.Append(justification14);
            paragraphProperties15.Append(paragraphMarkRunProperties15);

            paragraph18.Append(paragraphProperties15);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            Justification justification15 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            FontSize fontSize38 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties16.Append(fontSize38);
            paragraphMarkRunProperties16.Append(fontSizeComplexScript37);

            paragraphProperties16.Append(justification15);
            paragraphProperties16.Append(paragraphMarkRunProperties16);

            paragraph19.Append(paragraphProperties16);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();
            Justification justification16 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            FontSize fontSize39 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties17.Append(fontSize39);
            paragraphMarkRunProperties17.Append(fontSizeComplexScript38);

            paragraphProperties17.Append(justification16);
            paragraphProperties17.Append(paragraphMarkRunProperties17);
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

            paragraph20.Append(paragraphProperties17);
            paragraph20.Append(bookmarkStart1);
            paragraph20.Append(bookmarkEnd1);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph12);
            tableCell6.Append(paragraph13);
            tableCell6.Append(paragraph14);
            tableCell6.Append(paragraph15);
            tableCell6.Append(paragraph16);
            tableCell6.Append(paragraph17);
            tableCell6.Append(paragraph18);
            tableCell6.Append(paragraph19);
            tableCell6.Append(paragraph20);

            tableRow2.Append(tableRowProperties2);
            tableRow2.Append(tableCell4);
            tableRow2.Append(tableCell5);
            tableRow2.Append(tableCell6);

            TableRow tableRow3 = new TableRow() { RsidTableRowAddition = "008A46A1", RsidTableRowProperties = "004D6D95" };

            TableRowProperties tableRowProperties3 = new TableRowProperties();
            TableJustification tableJustification4 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            tableRowProperties3.Append(tableJustification4);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "4242", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan3 = new GridSpan() { Val = 4 };

            TableCellBorders tableCellBorders2 = new TableCellBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "0099FF", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(topBorder1);

            tableCellProperties7.Append(tableCellWidth7);
            tableCellProperties7.Append(gridSpan3);
            tableCellProperties7.Append(tableCellBorders2);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE9 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN9 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent9 = new AdjustRightIndent() { Val = false };
            Justification justification17 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            RunFonts runFonts27 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color26 = new Color() { Val = "231F20" };
            FontSize fontSize40 = new FontSize() { Val = "6" };
            FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "14" };

            paragraphMarkRunProperties18.Append(runFonts27);
            paragraphMarkRunProperties18.Append(color26);
            paragraphMarkRunProperties18.Append(fontSize40);
            paragraphMarkRunProperties18.Append(fontSizeComplexScript39);

            paragraphProperties18.Append(autoSpaceDE9);
            paragraphProperties18.Append(autoSpaceDN9);
            paragraphProperties18.Append(adjustRightIndent9);
            paragraphProperties18.Append(justification17);
            paragraphProperties18.Append(paragraphMarkRunProperties18);

            paragraph21.Append(paragraphProperties18);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE10 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN10 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent10 = new AdjustRightIndent() { Val = false };
            Justification justification18 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            RunFonts runFonts28 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color27 = new Color() { Val = "231F20" };
            FontSize fontSize41 = new FontSize() { Val = "6" };
            FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "14" };

            paragraphMarkRunProperties19.Append(runFonts28);
            paragraphMarkRunProperties19.Append(color27);
            paragraphMarkRunProperties19.Append(fontSize41);
            paragraphMarkRunProperties19.Append(fontSizeComplexScript40);

            paragraphProperties19.Append(autoSpaceDE10);
            paragraphProperties19.Append(autoSpaceDN10);
            paragraphProperties19.Append(adjustRightIndent10);
            paragraphProperties19.Append(justification18);
            paragraphProperties19.Append(paragraphMarkRunProperties19);

            paragraph22.Append(paragraphProperties19);

            Paragraph paragraph23 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE11 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN11 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent11 = new AdjustRightIndent() { Val = false };
            Justification justification19 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            RunFonts runFonts29 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color28 = new Color() { Val = "231F20" };
            FontSize fontSize42 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "14" };

            paragraphMarkRunProperties20.Append(runFonts29);
            paragraphMarkRunProperties20.Append(color28);
            paragraphMarkRunProperties20.Append(fontSize42);
            paragraphMarkRunProperties20.Append(fontSizeComplexScript41);

            paragraphProperties20.Append(autoSpaceDE11);
            paragraphProperties20.Append(autoSpaceDN11);
            paragraphProperties20.Append(adjustRightIndent11);
            paragraphProperties20.Append(justification19);
            paragraphProperties20.Append(paragraphMarkRunProperties20);
            ProofError proofError9 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run24 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties24 = new RunProperties();
            RunFonts runFonts30 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color29 = new Color() { Val = "231F20" };
            FontSize fontSize43 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "14" };

            runProperties24.Append(runFonts30);
            runProperties24.Append(color29);
            runProperties24.Append(fontSize43);
            runProperties24.Append(fontSizeComplexScript42);
            Text text23 = new Text();
            text23.Text = "Промзона";

            run24.Append(runProperties24);
            run24.Append(text23);
            ProofError proofError10 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run25 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties25 = new RunProperties();
            RunFonts runFonts31 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color30 = new Color() { Val = "231F20" };
            FontSize fontSize44 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "14" };

            runProperties25.Append(runFonts31);
            runProperties25.Append(color30);
            runProperties25.Append(fontSize44);
            runProperties25.Append(fontSizeComplexScript43);
            Text text24 = new Text();
            text24.Text = ", Панель №9,г";

            run25.Append(runProperties25);
            run25.Append(text24);

            Run run26 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties26 = new RunProperties();
            RunFonts runFonts32 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color31 = new Color() { Val = "231F20" };
            FontSize fontSize45 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "14" };

            runProperties26.Append(runFonts32);
            runProperties26.Append(color31);
            runProperties26.Append(fontSize45);
            runProperties26.Append(fontSizeComplexScript44);
            Text text25 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text25.Text = ". ";

            run26.Append(runProperties26);
            run26.Append(text25);

            Run run27 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties27 = new RunProperties();
            RunFonts runFonts33 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color32 = new Color() { Val = "231F20" };
            FontSize fontSize46 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "14" };

            runProperties27.Append(runFonts33);
            runProperties27.Append(color32);
            runProperties27.Append(fontSize46);
            runProperties27.Append(fontSizeComplexScript45);
            Text text26 = new Text();
            text26.Text = "Ноябрьск";

            run27.Append(runProperties27);
            run27.Append(text26);

            Run run28 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties28 = new RunProperties();
            RunFonts runFonts34 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color33 = new Color() { Val = "231F20" };
            FontSize fontSize47 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "14" };

            runProperties28.Append(runFonts34);
            runProperties28.Append(color33);
            runProperties28.Append(fontSize47);
            runProperties28.Append(fontSizeComplexScript46);
            Text text27 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text27.Text = ", ";

            run28.Append(runProperties28);
            run28.Append(text27);

            Run run29 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties29 = new RunProperties();
            RunFonts runFonts35 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color34 = new Color() { Val = "231F20" };
            FontSize fontSize48 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "14" };

            runProperties29.Append(runFonts35);
            runProperties29.Append(color34);
            runProperties29.Append(fontSize48);
            runProperties29.Append(fontSizeComplexScript47);
            Text text28 = new Text();
            text28.Text = "Ямало-Ненецкий";

            run29.Append(runProperties29);
            run29.Append(text28);

            paragraph23.Append(paragraphProperties20);
            paragraph23.Append(proofError9);
            paragraph23.Append(run24);
            paragraph23.Append(proofError10);
            paragraph23.Append(run25);
            paragraph23.Append(run26);
            paragraph23.Append(run27);
            paragraph23.Append(run28);
            paragraph23.Append(run29);

            Paragraph paragraph24 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE12 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN12 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent12 = new AdjustRightIndent() { Val = false };
            Justification justification20 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            RunFonts runFonts36 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color35 = new Color() { Val = "231F20" };
            FontSize fontSize49 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "14" };

            paragraphMarkRunProperties21.Append(runFonts36);
            paragraphMarkRunProperties21.Append(color35);
            paragraphMarkRunProperties21.Append(fontSize49);
            paragraphMarkRunProperties21.Append(fontSizeComplexScript48);

            paragraphProperties21.Append(autoSpaceDE12);
            paragraphProperties21.Append(autoSpaceDN12);
            paragraphProperties21.Append(adjustRightIndent12);
            paragraphProperties21.Append(justification20);
            paragraphProperties21.Append(paragraphMarkRunProperties21);
            ProofError proofError11 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run30 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties30 = new RunProperties();
            RunFonts runFonts37 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color36 = new Color() { Val = "231F20" };
            FontSize fontSize50 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "14" };

            runProperties30.Append(runFonts37);
            runProperties30.Append(color36);
            runProperties30.Append(fontSize50);
            runProperties30.Append(fontSizeComplexScript49);
            Text text29 = new Text();
            text29.Text = "автономныйокруг";

            run30.Append(runProperties30);
            run30.Append(text29);
            ProofError proofError12 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run31 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties31 = new RunProperties();
            RunFonts runFonts38 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color37 = new Color() { Val = "231F20" };
            FontSize fontSize51 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "14" };

            runProperties31.Append(runFonts38);
            runProperties31.Append(color37);
            runProperties31.Append(fontSize51);
            runProperties31.Append(fontSizeComplexScript50);
            Text text30 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text30.Text = ", ";

            run31.Append(runProperties31);
            run31.Append(text30);
            ProofError proofError13 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run32 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties32 = new RunProperties();
            RunFonts runFonts39 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color38 = new Color() { Val = "231F20" };
            FontSize fontSize52 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "14" };

            runProperties32.Append(runFonts39);
            runProperties32.Append(color38);
            runProperties32.Append(fontSize52);
            runProperties32.Append(fontSizeComplexScript51);
            Text text31 = new Text();
            text31.Text = "РоссийскаяФедерация";

            run32.Append(runProperties32);
            run32.Append(text31);
            ProofError proofError14 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run33 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties33 = new RunProperties();
            RunFonts runFonts40 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color39 = new Color() { Val = "231F20" };
            FontSize fontSize53 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript52 = new FontSizeComplexScript() { Val = "14" };

            runProperties33.Append(runFonts40);
            runProperties33.Append(color39);
            runProperties33.Append(fontSize53);
            runProperties33.Append(fontSizeComplexScript52);
            Text text32 = new Text();
            text32.Text = ", 629800";

            run33.Append(runProperties33);
            run33.Append(text32);

            paragraph24.Append(paragraphProperties21);
            paragraph24.Append(proofError11);
            paragraph24.Append(run30);
            paragraph24.Append(proofError12);
            paragraph24.Append(run31);
            paragraph24.Append(proofError13);
            paragraph24.Append(run32);
            paragraph24.Append(proofError14);
            paragraph24.Append(run33);

            Paragraph paragraph25 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE13 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN13 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent13 = new AdjustRightIndent() { Val = false };
            Justification justification21 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts41 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color40 = new Color() { Val = "231F20" };
            FontSize fontSize54 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript53 = new FontSizeComplexScript() { Val = "14" };

            paragraphMarkRunProperties22.Append(runFonts41);
            paragraphMarkRunProperties22.Append(color40);
            paragraphMarkRunProperties22.Append(fontSize54);
            paragraphMarkRunProperties22.Append(fontSizeComplexScript53);

            paragraphProperties22.Append(autoSpaceDE13);
            paragraphProperties22.Append(autoSpaceDN13);
            paragraphProperties22.Append(adjustRightIndent13);
            paragraphProperties22.Append(justification21);
            paragraphProperties22.Append(paragraphMarkRunProperties22);

            Run run34 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties34 = new RunProperties();
            RunFonts runFonts42 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color41 = new Color() { Val = "231F20" };
            FontSize fontSize55 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript54 = new FontSizeComplexScript() { Val = "14" };

            runProperties34.Append(runFonts42);
            runProperties34.Append(color41);
            runProperties34.Append(fontSize55);
            runProperties34.Append(fontSizeComplexScript54);
            Text text33 = new Text();
            text33.Text = "Тел";

            run34.Append(runProperties34);
            run34.Append(text33);

            Run run35 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties35 = new RunProperties();
            RunFonts runFonts43 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color42 = new Color() { Val = "231F20" };
            FontSize fontSize56 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript55 = new FontSizeComplexScript() { Val = "14" };

            runProperties35.Append(runFonts43);
            runProperties35.Append(color42);
            runProperties35.Append(fontSize56);
            runProperties35.Append(fontSizeComplexScript55);
            Text text34 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text34.Text = ".: (3496) 36-08-59, ";

            run35.Append(runProperties35);
            run35.Append(text34);

            Run run36 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties36 = new RunProperties();
            RunFonts runFonts44 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color43 = new Color() { Val = "231F20" };
            FontSize fontSize57 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript56 = new FontSizeComplexScript() { Val = "14" };

            runProperties36.Append(runFonts44);
            runProperties36.Append(color43);
            runProperties36.Append(fontSize57);
            runProperties36.Append(fontSizeComplexScript56);
            Text text35 = new Text();
            text35.Text = "факс";

            run36.Append(runProperties36);
            run36.Append(text35);

            Run run37 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties37 = new RunProperties();
            RunFonts runFonts45 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color44 = new Color() { Val = "231F20" };
            FontSize fontSize58 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript57 = new FontSizeComplexScript() { Val = "14" };

            runProperties37.Append(runFonts45);
            runProperties37.Append(color44);
            runProperties37.Append(fontSize58);
            runProperties37.Append(fontSizeComplexScript57);
            Text text36 = new Text();
            text36.Text = ": (3496) 36-08-60";

            run37.Append(runProperties37);
            run37.Append(text36);

            paragraph25.Append(paragraphProperties22);
            paragraph25.Append(run34);
            paragraph25.Append(run35);
            paragraph25.Append(run36);
            paragraph25.Append(run37);

            Paragraph paragraph26 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE14 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN14 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent14 = new AdjustRightIndent() { Val = false };
            Justification justification22 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            RunFonts runFonts46 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color45 = new Color() { Val = "231F20" };
            FontSize fontSize59 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript58 = new FontSizeComplexScript() { Val = "14" };
            Languages languages9 = new Languages() { Val = "de-DE" };

            paragraphMarkRunProperties23.Append(runFonts46);
            paragraphMarkRunProperties23.Append(color45);
            paragraphMarkRunProperties23.Append(fontSize59);
            paragraphMarkRunProperties23.Append(fontSizeComplexScript58);
            paragraphMarkRunProperties23.Append(languages9);

            paragraphProperties23.Append(autoSpaceDE14);
            paragraphProperties23.Append(autoSpaceDN14);
            paragraphProperties23.Append(adjustRightIndent14);
            paragraphProperties23.Append(justification22);
            paragraphProperties23.Append(paragraphMarkRunProperties23);

            Run run38 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties38 = new RunProperties();
            Color color46 = new Color() { Val = "231F20" };
            FontSize fontSize60 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript59 = new FontSizeComplexScript() { Val = "14" };
            Languages languages10 = new Languages() { Val = "de-DE" };

            runProperties38.Append(color46);
            runProperties38.Append(fontSize60);
            runProperties38.Append(fontSizeComplexScript59);
            runProperties38.Append(languages10);
            Text text37 = new Text();
            text37.Text = "E";

            run38.Append(runProperties38);
            run38.Append(text37);

            Run run39 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties39 = new RunProperties();
            RunFonts runFonts47 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color47 = new Color() { Val = "231F20" };
            FontSize fontSize61 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript60 = new FontSizeComplexScript() { Val = "14" };
            Languages languages11 = new Languages() { Val = "de-DE" };

            runProperties39.Append(runFonts47);
            runProperties39.Append(color47);
            runProperties39.Append(fontSize61);
            runProperties39.Append(fontSizeComplexScript60);
            runProperties39.Append(languages11);
            Text text38 = new Text();
            text38.Text = "-";

            run39.Append(runProperties39);
            run39.Append(text38);
            ProofError proofError15 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run40 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties40 = new RunProperties();
            RunFonts runFonts48 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color48 = new Color() { Val = "231F20" };
            FontSize fontSize62 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript61 = new FontSizeComplexScript() { Val = "14" };
            Languages languages12 = new Languages() { Val = "de-DE" };

            runProperties40.Append(runFonts48);
            runProperties40.Append(color48);
            runProperties40.Append(fontSize62);
            runProperties40.Append(fontSizeComplexScript61);
            runProperties40.Append(languages12);
            Text text39 = new Text();
            text39.Text = "mail:";

            run40.Append(runProperties40);
            run40.Append(text39);

            Run run41 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties41 = new RunProperties();
            RunFonts runFonts49 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "Courier New CYR" };
            FontSize fontSize63 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript62 = new FontSizeComplexScript() { Val = "14" };
            Languages languages13 = new Languages() { Val = "de-DE" };

            runProperties41.Append(runFonts49);
            runProperties41.Append(fontSize63);
            runProperties41.Append(fontSizeComplexScript62);
            runProperties41.Append(languages13);
            Text text40 = new Text();
            text40.Text = "info@noyabrsk";

            run41.Append(runProperties41);
            run41.Append(text40);
            ProofError proofError16 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run42 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties42 = new RunProperties();
            RunFonts runFonts50 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "Courier New CYR" };
            FontSize fontSize64 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript63 = new FontSizeComplexScript() { Val = "14" };
            Languages languages14 = new Languages() { Val = "de-DE" };

            runProperties42.Append(runFonts50);
            runProperties42.Append(fontSize64);
            runProperties42.Append(fontSizeComplexScript63);
            runProperties42.Append(languages14);
            Text text41 = new Text();
            text41.Text = "-";

            run42.Append(runProperties42);
            run42.Append(text41);
            ProofError proofError17 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run43 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties43 = new RunProperties();
            RunFonts runFonts51 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "Courier New CYR" };
            FontSize fontSize65 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript64 = new FontSizeComplexScript() { Val = "14" };
            Languages languages15 = new Languages() { Val = "de-DE" };

            runProperties43.Append(runFonts51);
            runProperties43.Append(fontSize65);
            runProperties43.Append(fontSizeComplexScript64);
            runProperties43.Append(languages15);
            Text text42 = new Text();
            text42.Text = "dobycha.gazprom.ru";

            run43.Append(runProperties43);
            run43.Append(text42);
            ProofError proofError18 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run44 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties44 = new RunProperties();
            RunFonts runFonts52 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color49 = new Color() { Val = "231F20" };
            FontSize fontSize66 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript65 = new FontSizeComplexScript() { Val = "14" };
            Languages languages16 = new Languages() { Val = "de-DE" };

            runProperties44.Append(runFonts52);
            runProperties44.Append(color49);
            runProperties44.Append(fontSize66);
            runProperties44.Append(fontSizeComplexScript65);
            runProperties44.Append(languages16);
            Text text43 = new Text();
            text43.Text = ", www.gazprom.ru";

            run44.Append(runProperties44);
            run44.Append(text43);

            paragraph26.Append(paragraphProperties23);
            paragraph26.Append(run38);
            paragraph26.Append(run39);
            paragraph26.Append(proofError15);
            paragraph26.Append(run40);
            paragraph26.Append(run41);
            paragraph26.Append(proofError16);
            paragraph26.Append(run42);
            paragraph26.Append(proofError17);
            paragraph26.Append(run43);
            paragraph26.Append(proofError18);
            paragraph26.Append(run44);

            Paragraph paragraph27 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE15 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN15 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent15 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            RunFonts runFonts53 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color50 = new Color() { Val = "231F20" };
            FontSize fontSize67 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript66 = new FontSizeComplexScript() { Val = "14" };

            paragraphMarkRunProperties24.Append(runFonts53);
            paragraphMarkRunProperties24.Append(color50);
            paragraphMarkRunProperties24.Append(fontSize67);
            paragraphMarkRunProperties24.Append(fontSizeComplexScript66);

            paragraphProperties24.Append(autoSpaceDE15);
            paragraphProperties24.Append(autoSpaceDN15);
            paragraphProperties24.Append(adjustRightIndent15);
            paragraphProperties24.Append(paragraphMarkRunProperties24);

            Run run45 = new Run();

            RunProperties runProperties45 = new RunProperties();
            RunFonts runFonts54 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color51 = new Color() { Val = "231F20" };
            FontSize fontSize68 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript67 = new FontSizeComplexScript() { Val = "14" };

            runProperties45.Append(runFonts54);
            runProperties45.Append(color51);
            runProperties45.Append(fontSize68);
            runProperties45.Append(fontSizeComplexScript67);
            Text text44 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text44.Text = "                             ";

            run45.Append(runProperties45);
            run45.Append(text44);
            ProofError proofError19 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run46 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties46 = new RunProperties();
            RunFonts runFonts55 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color52 = new Color() { Val = "231F20" };
            FontSize fontSize69 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript68 = new FontSizeComplexScript() { Val = "14" };

            runProperties46.Append(runFonts55);
            runProperties46.Append(color52);
            runProperties46.Append(fontSize69);
            runProperties46.Append(fontSizeComplexScript68);
            Text text45 = new Text();
            text45.Text = "OK";

            run46.Append(runProperties46);
            run46.Append(text45);
            ProofError proofError20 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run47 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties47 = new RunProperties();
            RunFonts runFonts56 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color53 = new Color() { Val = "231F20" };
            FontSize fontSize70 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript69 = new FontSizeComplexScript() { Val = "14" };

            runProperties47.Append(runFonts56);
            runProperties47.Append(color53);
            runProperties47.Append(fontSize70);
            runProperties47.Append(fontSizeComplexScript69);
            Text text46 = new Text();
            text46.Text = "ПО 05751797";

            run47.Append(runProperties47);
            run47.Append(text46);

            Run run48 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties48 = new RunProperties();
            RunFonts runFonts57 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "HeliosCond" };
            Color color54 = new Color() { Val = "231F20" };
            FontSize fontSize71 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript70 = new FontSizeComplexScript() { Val = "14" };

            runProperties48.Append(runFonts57);
            runProperties48.Append(color54);
            runProperties48.Append(fontSize71);
            runProperties48.Append(fontSizeComplexScript70);
            Text text47 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text47.Text = ", ";

            run48.Append(runProperties48);
            run48.Append(text47);

            Run run49 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties49 = new RunProperties();
            RunFonts runFonts58 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color55 = new Color() { Val = "231F20" };
            FontSize fontSize72 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript71 = new FontSizeComplexScript() { Val = "14" };

            runProperties49.Append(runFonts58);
            runProperties49.Append(color55);
            runProperties49.Append(fontSize72);
            runProperties49.Append(fontSizeComplexScript71);
            Text text48 = new Text();
            text48.Text = "ОГРН 1028900706647";

            run49.Append(runProperties49);
            run49.Append(text48);

            paragraph27.Append(paragraphProperties24);
            paragraph27.Append(run45);
            paragraph27.Append(proofError19);
            paragraph27.Append(run46);
            paragraph27.Append(proofError20);
            paragraph27.Append(run47);
            paragraph27.Append(run48);
            paragraph27.Append(run49);

            Paragraph paragraph28 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE16 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN16 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent16 = new AdjustRightIndent() { Val = false };
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Justification justification23 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            RunFonts runFonts59 = new RunFonts() { ComplexScript = "HeliosCond" };
            Color color56 = new Color() { Val = "000000" };
            FontSize fontSize73 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript72 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties25.Append(runFonts59);
            paragraphMarkRunProperties25.Append(color56);
            paragraphMarkRunProperties25.Append(fontSize73);
            paragraphMarkRunProperties25.Append(fontSizeComplexScript72);

            paragraphProperties25.Append(autoSpaceDE16);
            paragraphProperties25.Append(autoSpaceDN16);
            paragraphProperties25.Append(adjustRightIndent16);
            paragraphProperties25.Append(spacingBetweenLines6);
            paragraphProperties25.Append(justification23);
            paragraphProperties25.Append(paragraphMarkRunProperties25);

            Run run50 = new Run() { RsidRunProperties = "00FE335F" };

            RunProperties runProperties50 = new RunProperties();
            RunFonts runFonts60 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond", ComplexScript = "Courier New CYR" };
            FontSize fontSize74 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript73 = new FontSizeComplexScript() { Val = "14" };
            Languages languages17 = new Languages() { Val = "de-DE" };

            runProperties50.Append(runFonts60);
            runProperties50.Append(fontSize74);
            runProperties50.Append(fontSizeComplexScript73);
            runProperties50.Append(languages17);
            Text text49 = new Text();
            text49.Text = "ИНН/КПП 8905026850/997250001";

            run50.Append(runProperties50);
            run50.Append(text49);

            paragraph28.Append(paragraphProperties25);
            paragraph28.Append(run50);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph21);
            tableCell7.Append(paragraph22);
            tableCell7.Append(paragraph23);
            tableCell7.Append(paragraph24);
            tableCell7.Append(paragraph25);
            tableCell7.Append(paragraph26);
            tableCell7.Append(paragraph27);
            tableCell7.Append(paragraph28);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "972", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge3 = new VerticalMerge();

            tableCellProperties8.Append(tableCellWidth8);
            tableCellProperties8.Append(verticalMerge3);
            Paragraph paragraph29 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph29);

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "4646", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge4 = new VerticalMerge();

            tableCellProperties9.Append(tableCellWidth9);
            tableCellProperties9.Append(verticalMerge4);
            Paragraph paragraph30 = new Paragraph() { RsidParagraphMarkRevision = "003516C1", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph30);

            tableRow3.Append(tableRowProperties3);
            tableRow3.Append(tableCell7);
            tableRow3.Append(tableCell8);
            tableRow3.Append(tableCell9);

            TableRow tableRow4 = new TableRow() { RsidTableRowAddition = "008A46A1", RsidTableRowProperties = "004D6D95" };

            TableRowProperties tableRowProperties4 = new TableRowProperties();
            TableJustification tableJustification5 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            tableRowProperties4.Append(tableJustification5);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "1955", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan4 = new GridSpan() { Val = 2 };

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders3.Append(bottomBorder2);

            TableCellMargin tableCellMargin1 = new TableCellMargin();
            LeftMargin leftMargin1 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin1 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin1.Append(leftMargin1);
            tableCellMargin1.Append(rightMargin1);
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties10.Append(tableCellWidth10);
            tableCellProperties10.Append(gridSpan4);
            tableCellProperties10.Append(tableCellBorders3);
            tableCellProperties10.Append(tableCellMargin1);
            tableCellProperties10.Append(tableCellVerticalAlignment2);

            Paragraph paragraph31 = new Paragraph() { RsidParagraphMarkRevision = "00BC2F1A", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "00CC3955" };

            ParagraphProperties paragraphProperties26 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE17 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN17 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent17 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties26 = new ParagraphMarkRunProperties();
            RunFonts runFonts61 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Bold bold20 = new Bold();
            Color color57 = new Color() { Val = "231F20" };
            FontSize fontSize75 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript74 = new FontSizeComplexScript() { Val = "16" };
            Languages languages18 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties26.Append(runFonts61);
            paragraphMarkRunProperties26.Append(bold20);
            paragraphMarkRunProperties26.Append(color57);
            paragraphMarkRunProperties26.Append(fontSize75);
            paragraphMarkRunProperties26.Append(fontSizeComplexScript74);
            paragraphMarkRunProperties26.Append(languages18);

            paragraphProperties26.Append(autoSpaceDE17);
            paragraphProperties26.Append(autoSpaceDN17);
            paragraphProperties26.Append(adjustRightIndent17);
            paragraphProperties26.Append(paragraphMarkRunProperties26);
            ProofError proofError21 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run51 = new Run();

            RunProperties runProperties51 = new RunProperties();
            RunFonts runFonts62 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Bold bold21 = new Bold();
            Color color58 = new Color() { Val = "231F20" };
            FontSize fontSize76 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript75 = new FontSizeComplexScript() { Val = "16" };
            Languages languages19 = new Languages() { Val = "en-US" };

            runProperties51.Append(runFonts62);
            runProperties51.Append(bold21);
            runProperties51.Append(color58);
            runProperties51.Append(fontSize76);
            runProperties51.Append(fontSizeComplexScript75);
            runProperties51.Append(languages19);
            Text text50 = new Text();
            text50.Text = Template.OutboxDate;

            run51.Append(runProperties51);
            run51.Append(text50);
            ProofError proofError22 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph31.Append(paragraphProperties26);
            paragraph31.Append(proofError21);
            paragraph31.Append(run51);
            paragraph31.Append(proofError22);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph31);

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "590", Type = TableWidthUnitValues.Dxa };

            TableCellMargin tableCellMargin2 = new TableCellMargin();
            LeftMargin leftMargin2 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin2 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin2.Append(leftMargin2);
            tableCellMargin2.Append(rightMargin2);
            TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Bottom };

            tableCellProperties11.Append(tableCellWidth11);
            tableCellProperties11.Append(tableCellMargin2);
            tableCellProperties11.Append(tableCellVerticalAlignment3);

            Paragraph paragraph32 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE18 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN18 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent18 = new AdjustRightIndent() { Val = false };
            Justification justification24 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties27 = new ParagraphMarkRunProperties();
            RunFonts runFonts63 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color59 = new Color() { Val = "231F20" };
            FontSize fontSize77 = new FontSize() { Val = "6" };
            FontSizeComplexScript fontSizeComplexScript76 = new FontSizeComplexScript() { Val = "14" };

            paragraphMarkRunProperties27.Append(runFonts63);
            paragraphMarkRunProperties27.Append(color59);
            paragraphMarkRunProperties27.Append(fontSize77);
            paragraphMarkRunProperties27.Append(fontSizeComplexScript76);

            paragraphProperties27.Append(autoSpaceDE18);
            paragraphProperties27.Append(autoSpaceDN18);
            paragraphProperties27.Append(adjustRightIndent18);
            paragraphProperties27.Append(justification24);
            paragraphProperties27.Append(paragraphMarkRunProperties27);

            Run run52 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties52 = new RunProperties();
            RunFonts runFonts64 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color60 = new Color() { Val = "231F20" };
            FontSize fontSize78 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript77 = new FontSizeComplexScript() { Val = "18" };

            runProperties52.Append(runFonts64);
            runProperties52.Append(color60);
            runProperties52.Append(fontSize78);
            runProperties52.Append(fontSizeComplexScript77);
            Text text51 = new Text();
            text51.Text = "№";

            run52.Append(runProperties52);
            run52.Append(text51);

            paragraph32.Append(paragraphProperties27);
            paragraph32.Append(run52);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph32);

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "1697", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders4.Append(bottomBorder3);

            TableCellMargin tableCellMargin3 = new TableCellMargin();
            LeftMargin leftMargin3 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin3 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin3.Append(leftMargin3);
            tableCellMargin3.Append(rightMargin3);
            TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties12.Append(tableCellWidth12);
            tableCellProperties12.Append(tableCellBorders4);
            tableCellProperties12.Append(tableCellMargin3);
            tableCellProperties12.Append(tableCellVerticalAlignment4);

            Paragraph paragraph33 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "00CC3955" };

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE19 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN19 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent19 = new AdjustRightIndent() { Val = false };
            Justification justification25 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties28 = new ParagraphMarkRunProperties();
            RunFonts runFonts65 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Bold bold22 = new Bold();
            Color color61 = new Color() { Val = "231F20" };
            FontSize fontSize79 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript78 = new FontSizeComplexScript() { Val = "16" };
            Languages languages20 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties28.Append(runFonts65);
            paragraphMarkRunProperties28.Append(bold22);
            paragraphMarkRunProperties28.Append(color61);
            paragraphMarkRunProperties28.Append(fontSize79);
            paragraphMarkRunProperties28.Append(fontSizeComplexScript78);
            paragraphMarkRunProperties28.Append(languages20);

            paragraphProperties28.Append(autoSpaceDE19);
            paragraphProperties28.Append(autoSpaceDN19);
            paragraphProperties28.Append(adjustRightIndent19);
            paragraphProperties28.Append(justification25);
            paragraphProperties28.Append(paragraphMarkRunProperties28);
            ProofError proofError23 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run53 = new Run();

            RunProperties runProperties53 = new RunProperties();
            RunFonts runFonts66 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Bold bold23 = new Bold();
            Color color62 = new Color() { Val = "231F20" };
            FontSize fontSize80 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript79 = new FontSizeComplexScript() { Val = "16" };
            Languages languages21 = new Languages() { Val = "en-US" };

            runProperties53.Append(runFonts66);
            runProperties53.Append(bold23);
            runProperties53.Append(color62);
            runProperties53.Append(fontSize80);
            runProperties53.Append(fontSizeComplexScript79);
            runProperties53.Append(languages21);
            Text text52 = new Text();
            text52.Text = Template.OutboxNum;

            run53.Append(runProperties53);
            run53.Append(text52);
            ProofError proofError24 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph33.Append(paragraphProperties28);
            paragraph33.Append(proofError23);
            paragraph33.Append(run53);
            paragraph33.Append(proofError24);

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph33);

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "972", Type = TableWidthUnitValues.Dxa };

            tableCellProperties13.Append(tableCellWidth13);
            Paragraph paragraph34 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            tableCell13.Append(tableCellProperties13);
            tableCell13.Append(paragraph34);

            TableCell tableCell14 = new TableCell();

            TableCellProperties tableCellProperties14 = new TableCellProperties();
            TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "4646", Type = TableWidthUnitValues.Dxa };

            tableCellProperties14.Append(tableCellWidth14);
            Paragraph paragraph35 = new Paragraph() { RsidParagraphMarkRevision = "003516C1", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            tableCell14.Append(tableCellProperties14);
            tableCell14.Append(paragraph35);

            tableRow4.Append(tableRowProperties4);
            tableRow4.Append(tableCell10);
            tableRow4.Append(tableCell11);
            tableRow4.Append(tableCell12);
            tableRow4.Append(tableCell13);
            tableRow4.Append(tableCell14);

            TableRow tableRow5 = new TableRow() { RsidTableRowAddition = "008A46A1", RsidTableRowProperties = "004D6D95" };

            TableRowProperties tableRowProperties5 = new TableRowProperties();
            TableJustification tableJustification6 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            tableRowProperties5.Append(tableJustification6);

            TableCell tableCell15 = new TableCell();

            TableCellProperties tableCellProperties15 = new TableCellProperties();
            TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "538", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders5 = new TableCellBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders5.Append(topBorder2);

            TableCellMargin tableCellMargin4 = new TableCellMargin();
            LeftMargin leftMargin4 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin4 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin4.Append(leftMargin4);
            tableCellMargin4.Append(rightMargin4);
            TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Bottom };

            tableCellProperties15.Append(tableCellWidth15);
            tableCellProperties15.Append(tableCellBorders5);
            tableCellProperties15.Append(tableCellMargin4);
            tableCellProperties15.Append(tableCellVerticalAlignment5);

            Paragraph paragraph36 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE20 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN20 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent20 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties29 = new ParagraphMarkRunProperties();
            RunFonts runFonts67 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color63 = new Color() { Val = "231F20" };
            FontSize fontSize81 = new FontSize() { Val = "6" };
            FontSizeComplexScript fontSizeComplexScript80 = new FontSizeComplexScript() { Val = "14" };

            paragraphMarkRunProperties29.Append(runFonts67);
            paragraphMarkRunProperties29.Append(color63);
            paragraphMarkRunProperties29.Append(fontSize81);
            paragraphMarkRunProperties29.Append(fontSizeComplexScript80);

            paragraphProperties29.Append(autoSpaceDE20);
            paragraphProperties29.Append(autoSpaceDN20);
            paragraphProperties29.Append(adjustRightIndent20);
            paragraphProperties29.Append(paragraphMarkRunProperties29);
            ProofError proofError25 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run54 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties54 = new RunProperties();
            RunFonts runFonts68 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color64 = new Color() { Val = "231F20" };
            FontSize fontSize82 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript81 = new FontSizeComplexScript() { Val = "18" };

            runProperties54.Append(runFonts68);
            runProperties54.Append(color64);
            runProperties54.Append(fontSize82);
            runProperties54.Append(fontSizeComplexScript81);
            Text text53 = new Text();
            text53.Text = "на№";

            run54.Append(runProperties54);
            run54.Append(text53);
            ProofError proofError26 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph36.Append(paragraphProperties29);
            paragraph36.Append(proofError25);
            paragraph36.Append(run54);
            paragraph36.Append(proofError26);

            tableCell15.Append(tableCellProperties15);
            tableCell15.Append(paragraph36);

            TableCell tableCell16 = new TableCell();

            TableCellProperties tableCellProperties16 = new TableCellProperties();
            TableCellWidth tableCellWidth16 = new TableCellWidth() { Width = "1417", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders6 = new TableCellBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders6.Append(topBorder3);
            tableCellBorders6.Append(bottomBorder4);

            TableCellMargin tableCellMargin5 = new TableCellMargin();
            LeftMargin leftMargin5 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin5 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin5.Append(leftMargin5);
            tableCellMargin5.Append(rightMargin5);
            TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties16.Append(tableCellWidth16);
            tableCellProperties16.Append(tableCellBorders6);
            tableCellProperties16.Append(tableCellMargin5);
            tableCellProperties16.Append(tableCellVerticalAlignment6);

            Paragraph paragraph37 = new Paragraph() { RsidParagraphMarkRevision = "000B239B", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties30 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE21 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN21 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent21 = new AdjustRightIndent() { Val = false };
            Justification justification26 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties30 = new ParagraphMarkRunProperties();
            RunFonts runFonts69 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color65 = new Color() { Val = "231F20" };
            FontSize fontSize83 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript82 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties30.Append(runFonts69);
            paragraphMarkRunProperties30.Append(color65);
            paragraphMarkRunProperties30.Append(fontSize83);
            paragraphMarkRunProperties30.Append(fontSizeComplexScript82);

            paragraphProperties30.Append(autoSpaceDE21);
            paragraphProperties30.Append(autoSpaceDN21);
            paragraphProperties30.Append(adjustRightIndent21);
            paragraphProperties30.Append(justification26);
            paragraphProperties30.Append(paragraphMarkRunProperties30);

            paragraph37.Append(paragraphProperties30);

            tableCell16.Append(tableCellProperties16);
            tableCell16.Append(paragraph37);

            TableCell tableCell17 = new TableCell();

            TableCellProperties tableCellProperties17 = new TableCellProperties();
            TableCellWidth tableCellWidth17 = new TableCellWidth() { Width = "590", Type = TableWidthUnitValues.Dxa };

            TableCellMargin tableCellMargin6 = new TableCellMargin();
            LeftMargin leftMargin6 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin6 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin6.Append(leftMargin6);
            tableCellMargin6.Append(rightMargin6);
            TableCellVerticalAlignment tableCellVerticalAlignment7 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Bottom };

            tableCellProperties17.Append(tableCellWidth17);
            tableCellProperties17.Append(tableCellMargin6);
            tableCellProperties17.Append(tableCellVerticalAlignment7);

            Paragraph paragraph38 = new Paragraph() { RsidParagraphMarkRevision = "00DC080F", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties31 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE22 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN22 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent22 = new AdjustRightIndent() { Val = false };
            Justification justification27 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties31 = new ParagraphMarkRunProperties();
            RunFonts runFonts70 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color66 = new Color() { Val = "231F20" };
            FontSize fontSize84 = new FontSize() { Val = "6" };
            FontSizeComplexScript fontSizeComplexScript83 = new FontSizeComplexScript() { Val = "14" };

            paragraphMarkRunProperties31.Append(runFonts70);
            paragraphMarkRunProperties31.Append(color66);
            paragraphMarkRunProperties31.Append(fontSize84);
            paragraphMarkRunProperties31.Append(fontSizeComplexScript83);

            paragraphProperties31.Append(autoSpaceDE22);
            paragraphProperties31.Append(autoSpaceDN22);
            paragraphProperties31.Append(adjustRightIndent22);
            paragraphProperties31.Append(justification27);
            paragraphProperties31.Append(paragraphMarkRunProperties31);

            Run run55 = new Run() { RsidRunProperties = "00DC080F" };

            RunProperties runProperties55 = new RunProperties();
            RunFonts runFonts71 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color67 = new Color() { Val = "231F20" };
            FontSize fontSize85 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript84 = new FontSizeComplexScript() { Val = "18" };

            runProperties55.Append(runFonts71);
            runProperties55.Append(color67);
            runProperties55.Append(fontSize85);
            runProperties55.Append(fontSizeComplexScript84);
            Text text54 = new Text();
            text54.Text = "от";

            run55.Append(runProperties55);
            run55.Append(text54);

            paragraph38.Append(paragraphProperties31);
            paragraph38.Append(run55);

            tableCell17.Append(tableCellProperties17);
            tableCell17.Append(paragraph38);

            TableCell tableCell18 = new TableCell();

            TableCellProperties tableCellProperties18 = new TableCellProperties();
            TableCellWidth tableCellWidth18 = new TableCellWidth() { Width = "1697", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders7 = new TableCellBorders();
            TopBorder topBorder4 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder5 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders7.Append(topBorder4);
            tableCellBorders7.Append(bottomBorder5);

            TableCellMargin tableCellMargin7 = new TableCellMargin();
            LeftMargin leftMargin7 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin7 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin7.Append(leftMargin7);
            tableCellMargin7.Append(rightMargin7);
            TableCellVerticalAlignment tableCellVerticalAlignment8 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties18.Append(tableCellWidth18);
            tableCellProperties18.Append(tableCellBorders7);
            tableCellProperties18.Append(tableCellMargin7);
            tableCellProperties18.Append(tableCellVerticalAlignment8);

            Paragraph paragraph39 = new Paragraph() { RsidParagraphMarkRevision = "000B239B", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties32 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE23 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN23 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent23 = new AdjustRightIndent() { Val = false };
            Justification justification28 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties32 = new ParagraphMarkRunProperties();
            RunFonts runFonts72 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            Color color68 = new Color() { Val = "231F20" };
            FontSize fontSize86 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript85 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties32.Append(runFonts72);
            paragraphMarkRunProperties32.Append(color68);
            paragraphMarkRunProperties32.Append(fontSize86);
            paragraphMarkRunProperties32.Append(fontSizeComplexScript85);

            paragraphProperties32.Append(autoSpaceDE23);
            paragraphProperties32.Append(autoSpaceDN23);
            paragraphProperties32.Append(adjustRightIndent23);
            paragraphProperties32.Append(justification28);
            paragraphProperties32.Append(paragraphMarkRunProperties32);

            paragraph39.Append(paragraphProperties32);

            tableCell18.Append(tableCellProperties18);
            tableCell18.Append(paragraph39);

            TableCell tableCell19 = new TableCell();

            TableCellProperties tableCellProperties19 = new TableCellProperties();
            TableCellWidth tableCellWidth19 = new TableCellWidth() { Width = "972", Type = TableWidthUnitValues.Dxa };

            tableCellProperties19.Append(tableCellWidth19);
            Paragraph paragraph40 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            tableCell19.Append(tableCellProperties19);
            tableCell19.Append(paragraph40);

            TableCell tableCell20 = new TableCell();

            TableCellProperties tableCellProperties20 = new TableCellProperties();
            TableCellWidth tableCellWidth20 = new TableCellWidth() { Width = "4646", Type = TableWidthUnitValues.Dxa };

            tableCellProperties20.Append(tableCellWidth20);
            Paragraph paragraph41 = new Paragraph() { RsidParagraphMarkRevision = "003516C1", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "004D6D95", RsidRunAdditionDefault = "008A46A1" };

            tableCell20.Append(tableCellProperties20);
            tableCell20.Append(paragraph41);

            tableRow5.Append(tableRowProperties5);
            tableRow5.Append(tableCell15);
            tableRow5.Append(tableCell16);
            tableRow5.Append(tableCell17);
            tableRow5.Append(tableCell18);
            tableRow5.Append(tableCell19);
            tableRow5.Append(tableCell20);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);
            table1.Append(tableRow4);
            table1.Append(tableRow5);

            Paragraph paragraph42 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties33 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties33 = new ParagraphMarkRunProperties();
            RunStyle runStyle1 = new RunStyle() { Val = "a4" };
            RunFonts runFonts73 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Italic italic1 = new Italic();
            FontSize fontSize87 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript86 = new FontSizeComplexScript() { Val = "20" };
            Languages languages22 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties33.Append(runStyle1);
            paragraphMarkRunProperties33.Append(runFonts73);
            paragraphMarkRunProperties33.Append(italic1);
            paragraphMarkRunProperties33.Append(fontSize87);
            paragraphMarkRunProperties33.Append(fontSizeComplexScript86);
            paragraphMarkRunProperties33.Append(languages22);

            paragraphProperties33.Append(paragraphMarkRunProperties33);

            paragraph42.Append(paragraphProperties33);

            Paragraph paragraph43 = new Paragraph() { RsidParagraphMarkRevision = "00BB50A2", RsidParagraphAddition = "00CC3955", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "00CC3955" };

            ParagraphProperties paragraphProperties34 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties34 = new ParagraphMarkRunProperties();
            RunStyle runStyle2 = new RunStyle() { Val = "a4" };
            RunFonts runFonts74 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Italic italic2 = new Italic();
            FontSize fontSize88 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript87 = new FontSizeComplexScript() { Val = "20" };
            Languages languages23 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties34.Append(runStyle2);
            paragraphMarkRunProperties34.Append(runFonts74);
            paragraphMarkRunProperties34.Append(italic2);
            paragraphMarkRunProperties34.Append(fontSize88);
            paragraphMarkRunProperties34.Append(fontSizeComplexScript87);
            paragraphMarkRunProperties34.Append(languages23);

            paragraphProperties34.Append(paragraphMarkRunProperties34);
            ProofError proofError27 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run56 = new Run();

            RunProperties runProperties56 = new RunProperties();
            RunStyle runStyle3 = new RunStyle() { Val = "a4" };
            RunFonts runFonts75 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Italic italic3 = new Italic();
            FontSize fontSize89 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript88 = new FontSizeComplexScript() { Val = "20" };
            Languages languages24 = new Languages() { Val = "en-US" };

            runProperties56.Append(runStyle3);
            runProperties56.Append(runFonts75);
            runProperties56.Append(italic3);
            runProperties56.Append(fontSize89);
            runProperties56.Append(fontSizeComplexScript88);
            runProperties56.Append(languages24);
            Text text55 = new Text();
            text55.Text = Template.OutboxTheme;

            run56.Append(runProperties56);
            run56.Append(text55);
            ProofError proofError28 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph43.Append(paragraphProperties34);
            paragraph43.Append(proofError27);
            paragraph43.Append(run56);
            paragraph43.Append(proofError28);

            Paragraph paragraph44 = new Paragraph() { RsidParagraphMarkRevision = "00574A0E", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();
            Justification justification29 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties35 = new ParagraphMarkRunProperties();
            Bold bold24 = new Bold();
            FontSize fontSize90 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript89 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties35.Append(bold24);
            paragraphMarkRunProperties35.Append(fontSize90);
            paragraphMarkRunProperties35.Append(fontSizeComplexScript89);

            paragraphProperties35.Append(justification29);
            paragraphProperties35.Append(paragraphMarkRunProperties35);

            paragraph44.Append(paragraphProperties35);

            Paragraph paragraph45 = new Paragraph() { RsidParagraphMarkRevision = "00BB50A2", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "00CC3955" };

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();
            Justification justification30 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties36 = new ParagraphMarkRunProperties();
            Bold bold25 = new Bold();
            FontSize fontSize91 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript90 = new FontSizeComplexScript() { Val = "28" };
            Languages languages25 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties36.Append(bold25);
            paragraphMarkRunProperties36.Append(fontSize91);
            paragraphMarkRunProperties36.Append(fontSizeComplexScript90);
            paragraphMarkRunProperties36.Append(languages25);

            paragraphProperties36.Append(justification30);
            paragraphProperties36.Append(paragraphMarkRunProperties36);
            ProofError proofError29 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run57 = new Run();

            RunProperties runProperties57 = new RunProperties();
            Bold bold26 = new Bold();
            FontSize fontSize92 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript91 = new FontSizeComplexScript() { Val = "28" };
            Languages languages26 = new Languages() { Val = "en-US" };

            runProperties57.Append(bold26);
            runProperties57.Append(fontSize92);
            runProperties57.Append(fontSizeComplexScript91);
            runProperties57.Append(languages26);
            Text text56 = new Text();
            text56.Text = Template.DearReciever;

            run57.Append(runProperties57);
            run57.Append(text56);
            ProofError proofError30 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph45.Append(paragraphProperties36);
            paragraph45.Append(proofError29);
            paragraph45.Append(run57);
            paragraph45.Append(proofError30);

            Paragraph paragraph46 = new Paragraph() { RsidParagraphMarkRevision = "00615F2E", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties37 = new ParagraphMarkRunProperties();
            FontSize fontSize93 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript92 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties37.Append(fontSize93);
            paragraphMarkRunProperties37.Append(fontSizeComplexScript92);

            paragraphProperties37.Append(paragraphMarkRunProperties37);

            paragraph46.Append(paragraphProperties37);

            Paragraph paragraph47 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties38 = new ParagraphMarkRunProperties();
            Bold bold27 = new Bold();
            FontSize fontSize94 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties38.Append(bold27);
            paragraphMarkRunProperties38.Append(fontSize94);

            paragraphProperties38.Append(paragraphMarkRunProperties38);

            paragraph47.Append(paragraphProperties38);

            Paragraph paragraph48 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties39 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties39 = new ParagraphMarkRunProperties();
            Bold bold28 = new Bold();
            FontSize fontSize95 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties39.Append(bold28);
            paragraphMarkRunProperties39.Append(fontSize95);

            paragraphProperties39.Append(paragraphMarkRunProperties39);

            paragraph48.Append(paragraphProperties39);

            Paragraph paragraph49 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "00BB50A2" };

            ParagraphProperties paragraphProperties40 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties40 = new ParagraphMarkRunProperties();
            Bold bold29 = new Bold();
            FontSize fontSize96 = new FontSize() { Val = "28" };
            Languages languages27 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties40.Append(bold29);
            paragraphMarkRunProperties40.Append(fontSize96);
            paragraphMarkRunProperties40.Append(languages27);

            paragraphProperties40.Append(paragraphMarkRunProperties40);
            ProofError proofError31 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run58 = new Run();

            RunProperties runProperties58 = new RunProperties();
            Bold bold30 = new Bold();
            FontSize fontSize97 = new FontSize() { Val = "28" };
            Languages languages28 = new Languages() { Val = "en-US" };

            runProperties58.Append(bold30);
            runProperties58.Append(fontSize97);
            runProperties58.Append(languages28);
            Text text57 = new Text();
            text57.Text = Template.WhoSignPost;

            run58.Append(runProperties58);
            run58.Append(text57);
            ProofError proofError32 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run59 = new Run() { RsidRunAddition = "008A46A1" };

            RunProperties runProperties59 = new RunProperties();
            Bold bold31 = new Bold();
            FontSize fontSize98 = new FontSize() { Val = "28" };

            runProperties59.Append(bold31);
            runProperties59.Append(fontSize98);
            TabChar tabChar1 = new TabChar();

            run59.Append(runProperties59);
            run59.Append(tabChar1);

            Run run60 = new Run() { RsidRunAddition = "008A46A1" };

            RunProperties runProperties60 = new RunProperties();
            Bold bold32 = new Bold();
            FontSize fontSize99 = new FontSize() { Val = "28" };

            runProperties60.Append(bold32);
            runProperties60.Append(fontSize99);
            TabChar tabChar2 = new TabChar();

            run60.Append(runProperties60);
            run60.Append(tabChar2);

            Run run61 = new Run() { RsidRunAddition = "008A46A1" };

            RunProperties runProperties61 = new RunProperties();
            Bold bold33 = new Bold();
            FontSize fontSize100 = new FontSize() { Val = "28" };

            runProperties61.Append(bold33);
            runProperties61.Append(fontSize100);
            TabChar tabChar3 = new TabChar();

            run61.Append(runProperties61);
            run61.Append(tabChar3);

            Run run62 = new Run() { RsidRunAddition = "008A46A1" };

            RunProperties runProperties62 = new RunProperties();
            Bold bold34 = new Bold();
            FontSize fontSize101 = new FontSize() { Val = "28" };

            runProperties62.Append(bold34);
            runProperties62.Append(fontSize101);
            TabChar tabChar4 = new TabChar();
            Text text58 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text58.Text = " ";

            run62.Append(runProperties62);
            run62.Append(tabChar4);
            run62.Append(text58);
            ProofError proofError33 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run63 = new Run() { RsidRunAddition = "00CC3955" };

            RunProperties runProperties63 = new RunProperties();
            Bold bold35 = new Bold();
            FontSize fontSize102 = new FontSize() { Val = "28" };
            Languages languages29 = new Languages() { Val = "en-US" };

            runProperties63.Append(bold35);
            runProperties63.Append(fontSize102);
            runProperties63.Append(languages29);
            Text text59 = new Text();
            text59.Text = Template.WhoSignName;

            run63.Append(runProperties63);
            run63.Append(text59);

            Run run64 = new Run() { RsidRunAddition = "002B2A0B" };

            RunProperties runProperties64 = new RunProperties();
            Bold bold36 = new Bold();
            FontSize fontSize103 = new FontSize() { Val = "28" };
            Languages languages30 = new Languages() { Val = "en-US" };

            runProperties64.Append(bold36);
            runProperties64.Append(fontSize103);
            runProperties64.Append(languages30);
            Text text60 = new Text();
            text60.Text = "";

            run64.Append(runProperties64);
            run64.Append(text60);
            ProofError proofError34 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph49.Append(paragraphProperties40);
            paragraph49.Append(proofError31);
            paragraph49.Append(run58);
            paragraph49.Append(proofError32);
            paragraph49.Append(run59);
            paragraph49.Append(run60);
            paragraph49.Append(run61);
            paragraph49.Append(run62);
            paragraph49.Append(proofError33);
            paragraph49.Append(run63);
            paragraph49.Append(run64);
            paragraph49.Append(proofError34);

            Paragraph paragraph50 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties41 = new ParagraphProperties();
            Justification justification31 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties41 = new ParagraphMarkRunProperties();
            FontSize fontSize104 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript93 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties41.Append(fontSize104);
            paragraphMarkRunProperties41.Append(fontSizeComplexScript93);

            paragraphProperties41.Append(justification31);
            paragraphProperties41.Append(paragraphMarkRunProperties41);

            paragraph50.Append(paragraphProperties41);

            Paragraph paragraph51 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties42 = new ParagraphProperties();
            Justification justification32 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties42 = new ParagraphMarkRunProperties();
            FontSize fontSize105 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript94 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties42.Append(fontSize105);
            paragraphMarkRunProperties42.Append(fontSizeComplexScript94);

            paragraphProperties42.Append(justification32);
            paragraphProperties42.Append(paragraphMarkRunProperties42);

            paragraph51.Append(paragraphProperties42);

            Paragraph paragraph52 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties43 = new ParagraphProperties();
            Justification justification33 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties43 = new ParagraphMarkRunProperties();
            FontSize fontSize106 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript95 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties43.Append(fontSize106);
            paragraphMarkRunProperties43.Append(fontSizeComplexScript95);

            paragraphProperties43.Append(justification33);
            paragraphProperties43.Append(paragraphMarkRunProperties43);

            paragraph52.Append(paragraphProperties43);

            Paragraph paragraph53 = new Paragraph() { RsidParagraphMarkRevision = "006063B7", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties44 = new ParagraphProperties();
            Justification justification34 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties44 = new ParagraphMarkRunProperties();
            FontSize fontSize107 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript96 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties44.Append(fontSize107);
            paragraphMarkRunProperties44.Append(fontSizeComplexScript96);

            paragraphProperties44.Append(justification34);
            paragraphProperties44.Append(paragraphMarkRunProperties44);

            paragraph53.Append(paragraphProperties44);

            Paragraph paragraph54 = new Paragraph() { RsidParagraphMarkRevision = "006063B7", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties45 = new ParagraphProperties();
            Justification justification35 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties45 = new ParagraphMarkRunProperties();
            FontSize fontSize108 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript97 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties45.Append(fontSize108);
            paragraphMarkRunProperties45.Append(fontSizeComplexScript97);

            paragraphProperties45.Append(justification35);
            paragraphProperties45.Append(paragraphMarkRunProperties45);

            paragraph54.Append(paragraphProperties45);

            Paragraph paragraph55 = new Paragraph() { RsidParagraphMarkRevision = "006063B7", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties46 = new ParagraphProperties();
            Justification justification36 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties46 = new ParagraphMarkRunProperties();
            FontSize fontSize109 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript98 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties46.Append(fontSize109);
            paragraphMarkRunProperties46.Append(fontSizeComplexScript98);

            paragraphProperties46.Append(justification36);
            paragraphProperties46.Append(paragraphMarkRunProperties46);

            paragraph55.Append(paragraphProperties46);

            Paragraph paragraph56 = new Paragraph() { RsidParagraphMarkRevision = "006063B7", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties47 = new ParagraphProperties();
            Justification justification37 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties47 = new ParagraphMarkRunProperties();
            FontSize fontSize110 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript99 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties47.Append(fontSize110);
            paragraphMarkRunProperties47.Append(fontSizeComplexScript99);

            paragraphProperties47.Append(justification37);
            paragraphProperties47.Append(paragraphMarkRunProperties47);

            paragraph56.Append(paragraphProperties47);

            Paragraph paragraph57 = new Paragraph() { RsidParagraphMarkRevision = "006063B7", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties48 = new ParagraphProperties();
            Justification justification38 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties48 = new ParagraphMarkRunProperties();
            FontSize fontSize111 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript100 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties48.Append(fontSize111);
            paragraphMarkRunProperties48.Append(fontSizeComplexScript100);

            paragraphProperties48.Append(justification38);
            paragraphProperties48.Append(paragraphMarkRunProperties48);

            paragraph57.Append(paragraphProperties48);

            Paragraph paragraph58 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties49 = new ParagraphProperties();
            Justification justification39 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties49 = new ParagraphMarkRunProperties();
            FontSize fontSize112 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript101 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties49.Append(fontSize112);
            paragraphMarkRunProperties49.Append(fontSizeComplexScript101);

            paragraphProperties49.Append(justification39);
            paragraphProperties49.Append(paragraphMarkRunProperties49);

            paragraph58.Append(paragraphProperties49);

            Paragraph paragraph59 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties50 = new ParagraphProperties();
            Justification justification40 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties50 = new ParagraphMarkRunProperties();
            FontSize fontSize113 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript102 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties50.Append(fontSize113);
            paragraphMarkRunProperties50.Append(fontSizeComplexScript102);

            paragraphProperties50.Append(justification40);
            paragraphProperties50.Append(paragraphMarkRunProperties50);

            paragraph59.Append(paragraphProperties50);

            Paragraph paragraph60 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties51 = new ParagraphProperties();
            Justification justification41 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties51 = new ParagraphMarkRunProperties();
            FontSize fontSize114 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript103 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties51.Append(fontSize114);
            paragraphMarkRunProperties51.Append(fontSizeComplexScript103);

            paragraphProperties51.Append(justification41);
            paragraphProperties51.Append(paragraphMarkRunProperties51);

            paragraph60.Append(paragraphProperties51);

            Paragraph paragraph61 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties52 = new ParagraphProperties();
            Justification justification42 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties52 = new ParagraphMarkRunProperties();
            FontSize fontSize115 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript104 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties52.Append(fontSize115);
            paragraphMarkRunProperties52.Append(fontSizeComplexScript104);

            paragraphProperties52.Append(justification42);
            paragraphProperties52.Append(paragraphMarkRunProperties52);

            paragraph61.Append(paragraphProperties52);

            Paragraph paragraph62 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties53 = new ParagraphProperties();
            Justification justification43 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties53 = new ParagraphMarkRunProperties();
            FontSize fontSize116 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript105 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties53.Append(fontSize116);
            paragraphMarkRunProperties53.Append(fontSizeComplexScript105);

            paragraphProperties53.Append(justification43);
            paragraphProperties53.Append(paragraphMarkRunProperties53);

            paragraph62.Append(paragraphProperties53);

            Paragraph paragraph63 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties54 = new ParagraphProperties();
            Justification justification44 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties54 = new ParagraphMarkRunProperties();
            FontSize fontSize117 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript106 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties54.Append(fontSize117);
            paragraphMarkRunProperties54.Append(fontSizeComplexScript106);

            paragraphProperties54.Append(justification44);
            paragraphProperties54.Append(paragraphMarkRunProperties54);

            paragraph63.Append(paragraphProperties54);

            Paragraph paragraph64 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties55 = new ParagraphProperties();
            Justification justification45 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties55 = new ParagraphMarkRunProperties();
            FontSize fontSize118 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript107 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties55.Append(fontSize118);
            paragraphMarkRunProperties55.Append(fontSizeComplexScript107);

            paragraphProperties55.Append(justification45);
            paragraphProperties55.Append(paragraphMarkRunProperties55);

            paragraph64.Append(paragraphProperties55);

            Paragraph paragraph65 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties56 = new ParagraphProperties();
            Justification justification46 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties56 = new ParagraphMarkRunProperties();
            FontSize fontSize119 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript108 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties56.Append(fontSize119);
            paragraphMarkRunProperties56.Append(fontSizeComplexScript108);

            paragraphProperties56.Append(justification46);
            paragraphProperties56.Append(paragraphMarkRunProperties56);

            paragraph65.Append(paragraphProperties56);

            Paragraph paragraph66 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties57 = new ParagraphProperties();
            Justification justification47 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties57 = new ParagraphMarkRunProperties();
            FontSize fontSize120 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript109 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties57.Append(fontSize120);
            paragraphMarkRunProperties57.Append(fontSizeComplexScript109);

            paragraphProperties57.Append(justification47);
            paragraphProperties57.Append(paragraphMarkRunProperties57);

            paragraph66.Append(paragraphProperties57);

            Paragraph paragraph67 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties58 = new ParagraphProperties();
            Justification justification48 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties58 = new ParagraphMarkRunProperties();
            FontSize fontSize121 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript110 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties58.Append(fontSize121);
            paragraphMarkRunProperties58.Append(fontSizeComplexScript110);

            paragraphProperties58.Append(justification48);
            paragraphProperties58.Append(paragraphMarkRunProperties58);

            paragraph67.Append(paragraphProperties58);

            Paragraph paragraph68 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties59 = new ParagraphProperties();
            Justification justification49 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties59 = new ParagraphMarkRunProperties();
            FontSize fontSize122 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript111 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties59.Append(fontSize122);
            paragraphMarkRunProperties59.Append(fontSizeComplexScript111);

            paragraphProperties59.Append(justification49);
            paragraphProperties59.Append(paragraphMarkRunProperties59);

            paragraph68.Append(paragraphProperties59);

            Paragraph paragraph69 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties60 = new ParagraphProperties();
            Justification justification50 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties60 = new ParagraphMarkRunProperties();
            FontSize fontSize123 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript112 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties60.Append(fontSize123);
            paragraphMarkRunProperties60.Append(fontSizeComplexScript112);

            paragraphProperties60.Append(justification50);
            paragraphProperties60.Append(paragraphMarkRunProperties60);

            paragraph69.Append(paragraphProperties60);

            Paragraph paragraph70 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties61 = new ParagraphProperties();
            Justification justification51 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties61 = new ParagraphMarkRunProperties();
            FontSize fontSize124 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript113 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties61.Append(fontSize124);
            paragraphMarkRunProperties61.Append(fontSizeComplexScript113);

            paragraphProperties61.Append(justification51);
            paragraphProperties61.Append(paragraphMarkRunProperties61);

            paragraph70.Append(paragraphProperties61);

            Paragraph paragraph71 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties62 = new ParagraphProperties();
            Justification justification52 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties62 = new ParagraphMarkRunProperties();
            FontSize fontSize125 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript114 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties62.Append(fontSize125);
            paragraphMarkRunProperties62.Append(fontSizeComplexScript114);

            paragraphProperties62.Append(justification52);
            paragraphProperties62.Append(paragraphMarkRunProperties62);

            paragraph71.Append(paragraphProperties62);

            Paragraph paragraph72 = new Paragraph() { RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties63 = new ParagraphProperties();
            Justification justification53 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties63 = new ParagraphMarkRunProperties();
            FontSize fontSize126 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript115 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties63.Append(fontSize126);
            paragraphMarkRunProperties63.Append(fontSizeComplexScript115);

            paragraphProperties63.Append(justification53);
            paragraphProperties63.Append(paragraphMarkRunProperties63);

            paragraph72.Append(paragraphProperties63);

            Paragraph paragraph73 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "008A46A1" };

            ParagraphProperties paragraphProperties64 = new ParagraphProperties();
            Justification justification54 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties64 = new ParagraphMarkRunProperties();
            FontSize fontSize127 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript116 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties64.Append(fontSize127);
            paragraphMarkRunProperties64.Append(fontSizeComplexScript116);

            paragraphProperties64.Append(justification54);
            paragraphProperties64.Append(paragraphMarkRunProperties64);

            paragraph73.Append(paragraphProperties64);

            Paragraph paragraph74 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "00C71E0B", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "00C71E0B" };

            ParagraphProperties paragraphProperties65 = new ParagraphProperties();
            Justification justification55 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties65 = new ParagraphMarkRunProperties();
            FontSize fontSize128 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript117 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties65.Append(fontSize128);
            paragraphMarkRunProperties65.Append(fontSizeComplexScript117);

            paragraphProperties65.Append(justification55);
            paragraphProperties65.Append(paragraphMarkRunProperties65);

            paragraph74.Append(paragraphProperties65);

            Paragraph paragraph75 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "00C71E0B", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "00C71E0B" };

            ParagraphProperties paragraphProperties66 = new ParagraphProperties();
            Justification justification56 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties66 = new ParagraphMarkRunProperties();
            FontSize fontSize129 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript118 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties66.Append(fontSize129);
            paragraphMarkRunProperties66.Append(fontSizeComplexScript118);

            paragraphProperties66.Append(justification56);
            paragraphProperties66.Append(paragraphMarkRunProperties66);

            paragraph75.Append(paragraphProperties66);

            Paragraph paragraph76 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "00C71E0B", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "00C71E0B" };

            ParagraphProperties paragraphProperties67 = new ParagraphProperties();
            Justification justification57 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties67 = new ParagraphMarkRunProperties();
            FontSize fontSize130 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript119 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties67.Append(fontSize130);
            paragraphMarkRunProperties67.Append(fontSizeComplexScript119);

            paragraphProperties67.Append(justification57);
            paragraphProperties67.Append(paragraphMarkRunProperties67);

            paragraph76.Append(paragraphProperties67);

            Paragraph paragraph77 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "00C71E0B", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "00C71E0B" };

            ParagraphProperties paragraphProperties68 = new ParagraphProperties();
            Justification justification58 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties68 = new ParagraphMarkRunProperties();
            FontSize fontSize131 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript120 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties68.Append(fontSize131);
            paragraphMarkRunProperties68.Append(fontSizeComplexScript120);

            paragraphProperties68.Append(justification58);
            paragraphProperties68.Append(paragraphMarkRunProperties68);

            paragraph77.Append(paragraphProperties68);

            Paragraph paragraph78 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "00C71E0B", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "00C71E0B" };

            ParagraphProperties paragraphProperties69 = new ParagraphProperties();
            Justification justification59 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties69 = new ParagraphMarkRunProperties();
            FontSize fontSize132 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript121 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties69.Append(fontSize132);
            paragraphMarkRunProperties69.Append(fontSizeComplexScript121);

            paragraphProperties69.Append(justification59);
            paragraphProperties69.Append(paragraphMarkRunProperties69);

            paragraph78.Append(paragraphProperties69);

            Paragraph paragraph79 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "00C71E0B", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "00C71E0B" };

            ParagraphProperties paragraphProperties70 = new ParagraphProperties();
            Justification justification60 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties70 = new ParagraphMarkRunProperties();
            FontSize fontSize133 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript122 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties70.Append(fontSize133);
            paragraphMarkRunProperties70.Append(fontSizeComplexScript122);

            paragraphProperties70.Append(justification60);
            paragraphProperties70.Append(paragraphMarkRunProperties70);

            paragraph79.Append(paragraphProperties70);

            Paragraph paragraph80 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "00C71E0B", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "00C71E0B" };

            ParagraphProperties paragraphProperties71 = new ParagraphProperties();
            Justification justification61 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties71 = new ParagraphMarkRunProperties();
            FontSize fontSize134 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript123 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties71.Append(fontSize134);
            paragraphMarkRunProperties71.Append(fontSizeComplexScript123);

            paragraphProperties71.Append(justification61);
            paragraphProperties71.Append(paragraphMarkRunProperties71);

            paragraph80.Append(paragraphProperties71);

            Paragraph paragraph81 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "00C71E0B", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "00C71E0B" };

            ParagraphProperties paragraphProperties72 = new ParagraphProperties();
            Justification justification62 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties72 = new ParagraphMarkRunProperties();
            FontSize fontSize135 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript124 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties72.Append(fontSize135);
            paragraphMarkRunProperties72.Append(fontSizeComplexScript124);

            paragraphProperties72.Append(justification62);
            paragraphProperties72.Append(paragraphMarkRunProperties72);

            paragraph81.Append(paragraphProperties72);

            Paragraph paragraph82 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "00C71E0B", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "00C71E0B" };

            ParagraphProperties paragraphProperties73 = new ParagraphProperties();
            Justification justification63 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties73 = new ParagraphMarkRunProperties();
            FontSize fontSize136 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript125 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties73.Append(fontSize136);
            paragraphMarkRunProperties73.Append(fontSizeComplexScript125);

            paragraphProperties73.Append(justification63);
            paragraphProperties73.Append(paragraphMarkRunProperties73);

            paragraph82.Append(paragraphProperties73);

            Paragraph paragraph83 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "00C71E0B", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "00C71E0B" };

            ParagraphProperties paragraphProperties74 = new ParagraphProperties();
            Justification justification64 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties74 = new ParagraphMarkRunProperties();
            FontSize fontSize137 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript126 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties74.Append(fontSize137);
            paragraphMarkRunProperties74.Append(fontSizeComplexScript126);

            paragraphProperties74.Append(justification64);
            paragraphProperties74.Append(paragraphMarkRunProperties74);

            paragraph83.Append(paragraphProperties74);

            Paragraph paragraph84 = new Paragraph() { RsidParagraphMarkRevision = "00CC3955", RsidParagraphAddition = "00C71E0B", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "00C71E0B" };

            ParagraphProperties paragraphProperties75 = new ParagraphProperties();
            Justification justification65 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties75 = new ParagraphMarkRunProperties();
            FontSize fontSize138 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript127 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties75.Append(fontSize138);
            paragraphMarkRunProperties75.Append(fontSizeComplexScript127);

            paragraphProperties75.Append(justification65);
            paragraphProperties75.Append(paragraphMarkRunProperties75);

            paragraph84.Append(paragraphProperties75);

            Paragraph paragraph85 = new Paragraph() { RsidParagraphMarkRevision = "002B2A0B", RsidParagraphAddition = "008A46A1", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "002B2A0B" };

            ParagraphProperties paragraphProperties76 = new ParagraphProperties();
            Justification justification66 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties76 = new ParagraphMarkRunProperties();
            FontSize fontSize139 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript128 = new FontSizeComplexScript() { Val = "16" };
            Languages languages31 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties76.Append(fontSize139);
            paragraphMarkRunProperties76.Append(fontSizeComplexScript128);
            paragraphMarkRunProperties76.Append(languages31);

            paragraphProperties76.Append(justification66);
            paragraphProperties76.Append(paragraphMarkRunProperties76);
            ProofError proofError35 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run65 = new Run() { RsidRunProperties = "002B2A0B" };

            RunProperties runProperties65 = new RunProperties();
            FontSize fontSize140 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript129 = new FontSizeComplexScript() { Val = "16" };
            Languages languages32 = new Languages() { Val = "en-US" };

            runProperties65.Append(fontSize140);
            runProperties65.Append(fontSizeComplexScript129);
            runProperties65.Append(languages32);
            Text text61 = new Text();
            text61.Text = Template.WhoMadeName;

            run65.Append(runProperties65);
            run65.Append(text61);

            Run run66 = new Run();

            RunProperties runProperties66 = new RunProperties();
            FontSize fontSize141 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript130 = new FontSizeComplexScript() { Val = "16" };
            Languages languages33 = new Languages() { Val = "en-US" };

            runProperties66.Append(fontSize141);
            runProperties66.Append(fontSizeComplexScript130);
            runProperties66.Append(languages33);
            Text text62 = new Text();
            text62.Text = "";

            run66.Append(runProperties66);
            run66.Append(text62);
            ProofError proofError36 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph85.Append(paragraphProperties76);
            paragraph85.Append(proofError35);
            paragraph85.Append(run65);
            paragraph85.Append(run66);
            paragraph85.Append(proofError36);

            Paragraph paragraph86 = new Paragraph() { RsidParagraphMarkRevision = "002B2A0B", RsidParagraphAddition = "00576092", RsidParagraphProperties = "008A46A1", RsidRunAdditionDefault = "002B2A0B" };

            ParagraphProperties paragraphProperties77 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties77 = new ParagraphMarkRunProperties();
            FontSize fontSize142 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript131 = new FontSizeComplexScript() { Val = "16" };
            Languages languages34 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties77.Append(fontSize142);
            paragraphMarkRunProperties77.Append(fontSizeComplexScript131);
            paragraphMarkRunProperties77.Append(languages34);

            paragraphProperties77.Append(paragraphMarkRunProperties77);
            ProofError proofError37 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run67 = new Run() { RsidRunProperties = "002B2A0B" };

            RunProperties runProperties67 = new RunProperties();
            FontSize fontSize143 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript132 = new FontSizeComplexScript() { Val = "16" };
            Languages languages35 = new Languages() { Val = "en-US" };

            runProperties67.Append(fontSize143);
            runProperties67.Append(fontSizeComplexScript132);
            runProperties67.Append(languages35);
            Text text63 = new Text();
            text63.Text = Template.WhoMadeTel;

            run67.Append(runProperties67);
            run67.Append(text63);
            ProofError proofError38 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph86.Append(paragraphProperties77);
            paragraph86.Append(proofError37);
            paragraph86.Append(run67);
            paragraph86.Append(proofError38);

            SectionProperties sectionProperties1 = new SectionProperties() { RsidRPr = "002B2A0B", RsidR = "00576092", RsidSect = "009F2DB9" };
            FooterReference footerReference1 = new FooterReference() { Type = HeaderFooterValues.Even, Id = "rId9" };
            FooterReference footerReference2 = new FooterReference() { Type = HeaderFooterValues.Default, Id = "rId10" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U, Code = (UInt16Value)9U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1135, Right = (UInt32Value)851U, Bottom = 284, Left = (UInt32Value)1418U, Header = (UInt32Value)567U, Footer = (UInt32Value)567U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "708" };
            TitlePage titlePage1 = new TitlePage();
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(footerReference1);
            sectionProperties1.Append(footerReference2);
            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(titlePage1);
            sectionProperties1.Append(docGrid1);

            body1.Append(table1);
            body1.Append(paragraph42);
            body1.Append(paragraph43);
            body1.Append(paragraph44);
            body1.Append(paragraph45);
            body1.Append(paragraph46);
            body1.Append(paragraph47);
            body1.Append(paragraph48);
            body1.Append(paragraph49);
            body1.Append(paragraph50);
            body1.Append(paragraph51);
            body1.Append(paragraph52);
            body1.Append(paragraph53);
            body1.Append(paragraph54);
            body1.Append(paragraph55);
            body1.Append(paragraph56);
            body1.Append(paragraph57);
            body1.Append(paragraph58);
            body1.Append(paragraph59);
            body1.Append(paragraph60);
            body1.Append(paragraph61);
            body1.Append(paragraph62);
            body1.Append(paragraph63);
            body1.Append(paragraph64);
            body1.Append(paragraph65);
            body1.Append(paragraph66);
            body1.Append(paragraph67);
            body1.Append(paragraph68);
            body1.Append(paragraph69);
            body1.Append(paragraph70);
            body1.Append(paragraph71);
            body1.Append(paragraph72);
            body1.Append(paragraph73);
            body1.Append(paragraph74);
            body1.Append(paragraph75);
            body1.Append(paragraph76);
            body1.Append(paragraph77);
            body1.Append(paragraph78);
            body1.Append(paragraph79);
            body1.Append(paragraph80);
            body1.Append(paragraph81);
            body1.Append(paragraph82);
            body1.Append(paragraph83);
            body1.Append(paragraph84);
            body1.Append(paragraph85);
            body1.Append(paragraph86);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }

        // Generates content of imagePart1.
        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles();
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts76 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Languages languages36 = new Languages() { Val = "ru-RU", EastAsia = "ru-RU", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts76);
            runPropertiesBaseStyle1.Append(languages36);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);
            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 0, DefaultSemiHidden = false, DefaultUnhideWhenUsed = false, DefaultPrimaryStyle = false, Count = 267 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "caption", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "Title", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "Subtitle", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "Strong", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "Emphasis", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", UiPriority = 99, SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "Revision", UiPriority = 99, SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "a", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid1 = new Rsid() { Val = "008A46A1" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            Justification justification67 = new Justification() { Val = JustificationValues.Both };

            styleParagraphProperties1.Append(justification67);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            FontSize fontSize144 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript133 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties1.Append(fontSize144);
            styleRunProperties1.Append(fontSizeComplexScript133);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            Style style2 = new Style() { Type = StyleValues.Character, StyleId = "a0", Default = true };
            StyleName styleName2 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority1 = new UIPriority() { Val = 1 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            style2.Append(styleName2);
            style2.Append(uIPriority1);
            style2.Append(semiHidden1);
            style2.Append(unhideWhenUsed1);

            Style style3 = new Style() { Type = StyleValues.Table, StyleId = "a1", Default = true };
            StyleName styleName3 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority2 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle2 = new PrimaryStyle();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault2 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin2 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin2 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault2.Append(topMargin1);
            tableCellMarginDefault2.Append(tableCellLeftMargin2);
            tableCellMarginDefault2.Append(bottomMargin1);
            tableCellMarginDefault2.Append(tableCellRightMargin2);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault2);

            style3.Append(styleName3);
            style3.Append(uIPriority2);
            style3.Append(semiHidden2);
            style3.Append(unhideWhenUsed2);
            style3.Append(primaryStyle2);
            style3.Append(styleTableProperties1);

            Style style4 = new Style() { Type = StyleValues.Numbering, StyleId = "a2", Default = true };
            StyleName styleName4 = new StyleName() { Val = "No List" };
            UIPriority uIPriority3 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();

            style4.Append(styleName4);
            style4.Append(uIPriority3);
            style4.Append(semiHidden3);
            style4.Append(unhideWhenUsed3);

            Style style5 = new Style() { Type = StyleValues.Table, StyleId = "a3" };
            StyleName styleName5 = new StyleName() { Val = "Table Grid" };
            BasedOn basedOn1 = new BasedOn() { Val = "a1" };
            Rsid rsid2 = new Rsid() { Val = "007F5B6A" };

            StyleTableProperties styleTableProperties2 = new StyleTableProperties();
            TableIndentation tableIndentation2 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder5 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder6 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder5);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder6);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);

            TableCellMarginDefault tableCellMarginDefault3 = new TableCellMarginDefault();
            TopMargin topMargin2 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin3 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin2 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin3 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault3.Append(topMargin2);
            tableCellMarginDefault3.Append(tableCellLeftMargin3);
            tableCellMarginDefault3.Append(bottomMargin2);
            tableCellMarginDefault3.Append(tableCellRightMargin3);

            styleTableProperties2.Append(tableIndentation2);
            styleTableProperties2.Append(tableBorders1);
            styleTableProperties2.Append(tableCellMarginDefault3);

            style5.Append(styleName5);
            style5.Append(basedOn1);
            style5.Append(rsid2);
            style5.Append(styleTableProperties2);

            Style style6 = new Style() { Type = StyleValues.Character, StyleId = "a4", CustomStyle = true };
            StyleName styleName6 = new StyleName() { Val = "Стиль новый" };
            BasedOn basedOn2 = new BasedOn() { Val = "a0" };
            Rsid rsid3 = new Rsid() { Val = "00F512CC" };

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            RunFonts runFonts77 = new RunFonts() { Ascii = "HeliosCond", HighAnsi = "HeliosCond" };
            FontSize fontSize145 = new FontSize() { Val = "24" };

            styleRunProperties2.Append(runFonts77);
            styleRunProperties2.Append(fontSize145);

            style6.Append(styleName6);
            style6.Append(basedOn2);
            style6.Append(rsid3);
            style6.Append(styleRunProperties2);

            Style style7 = new Style() { Type = StyleValues.Paragraph, StyleId = "a5" };
            StyleName styleName7 = new StyleName() { Val = "footer" };
            BasedOn basedOn3 = new BasedOn() { Val = "a" };
            Rsid rsid4 = new Rsid() { Val = "00935E8D" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Center, Position = 4677 };
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Right, Position = 9355 };

            tabs1.Append(tabStop1);
            tabs1.Append(tabStop2);

            styleParagraphProperties2.Append(tabs1);

            style7.Append(styleName7);
            style7.Append(basedOn3);
            style7.Append(rsid4);
            style7.Append(styleParagraphProperties2);

            Style style8 = new Style() { Type = StyleValues.Character, StyleId = "a6" };
            StyleName styleName8 = new StyleName() { Val = "page number" };
            BasedOn basedOn4 = new BasedOn() { Val = "a0" };
            Rsid rsid5 = new Rsid() { Val = "00935E8D" };

            style8.Append(styleName8);
            style8.Append(basedOn4);
            style8.Append(rsid5);

            Style style9 = new Style() { Type = StyleValues.Paragraph, StyleId = "a7" };
            StyleName styleName9 = new StyleName() { Val = "header" };
            BasedOn basedOn5 = new BasedOn() { Val = "a" };
            Rsid rsid6 = new Rsid() { Val = "00935E8D" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();

            Tabs tabs2 = new Tabs();
            TabStop tabStop3 = new TabStop() { Val = TabStopValues.Center, Position = 4677 };
            TabStop tabStop4 = new TabStop() { Val = TabStopValues.Right, Position = 9355 };

            tabs2.Append(tabStop3);
            tabs2.Append(tabStop4);

            styleParagraphProperties3.Append(tabs2);

            style9.Append(styleName9);
            style9.Append(basedOn5);
            style9.Append(rsid6);
            style9.Append(styleParagraphProperties3);

            Style style10 = new Style() { Type = StyleValues.Paragraph, StyleId = "a8" };
            StyleName styleName10 = new StyleName() { Val = "Balloon Text" };
            BasedOn basedOn6 = new BasedOn() { Val = "a" };
            SemiHidden semiHidden4 = new SemiHidden();
            Rsid rsid7 = new Rsid() { Val = "00FD0BFD" };

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            RunFonts runFonts78 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            FontSize fontSize146 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript134 = new FontSizeComplexScript() { Val = "16" };

            styleRunProperties3.Append(runFonts78);
            styleRunProperties3.Append(fontSize146);
            styleRunProperties3.Append(fontSizeComplexScript134);

            style10.Append(styleName10);
            style10.Append(basedOn6);
            style10.Append(semiHidden4);
            style10.Append(rsid7);
            style10.Append(styleRunProperties3);

            Style style11 = new Style() { Type = StyleValues.Paragraph, StyleId = "a9" };
            StyleName styleName11 = new StyleName() { Val = "List" };
            BasedOn basedOn7 = new BasedOn() { Val = "a" };
            Rsid rsid8 = new Rsid() { Val = "00222D9A" };

            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();
            Indentation indentation1 = new Indentation() { Left = "283", Hanging = "283" };
            Justification justification68 = new Justification() { Val = JustificationValues.Left };

            styleParagraphProperties4.Append(indentation1);
            styleParagraphProperties4.Append(justification68);

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            FontSize fontSize147 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript135 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties4.Append(fontSize147);
            styleRunProperties4.Append(fontSizeComplexScript135);

            style11.Append(styleName11);
            style11.Append(basedOn7);
            style11.Append(rsid8);
            style11.Append(styleParagraphProperties4);
            style11.Append(styleRunProperties4);

            Style style12 = new Style() { Type = StyleValues.Paragraph, StyleId = "aa" };
            StyleName styleName12 = new StyleName() { Val = "List Paragraph" };
            BasedOn basedOn8 = new BasedOn() { Val = "a" };
            UIPriority uIPriority4 = new UIPriority() { Val = 34 };
            PrimaryStyle primaryStyle3 = new PrimaryStyle();
            Rsid rsid9 = new Rsid() { Val = "00C957A3" };

            StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();
            Indentation indentation2 = new Indentation() { Left = "720" };
            ContextualSpacing contextualSpacing1 = new ContextualSpacing();

            styleParagraphProperties5.Append(indentation2);
            styleParagraphProperties5.Append(contextualSpacing1);

            style12.Append(styleName12);
            style12.Append(basedOn8);
            style12.Append(uIPriority4);
            style12.Append(primaryStyle3);
            style12.Append(rsid9);
            style12.Append(styleParagraphProperties5);

            Style style13 = new Style() { Type = StyleValues.Character, StyleId = "ab" };
            StyleName styleName13 = new StyleName() { Val = "Placeholder Text" };
            BasedOn basedOn9 = new BasedOn() { Val = "a0" };
            UIPriority uIPriority5 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden5 = new SemiHidden();
            Rsid rsid10 = new Rsid() { Val = "00654D13" };

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            Color color69 = new Color() { Val = "808080" };

            styleRunProperties5.Append(color69);

            style13.Append(styleName13);
            style13.Append(basedOn9);
            style13.Append(uIPriority5);
            style13.Append(semiHidden5);
            style13.Append(rsid10);
            style13.Append(styleRunProperties5);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);
            styles1.Append(style8);
            styles1.Append(style9);
            styles1.Append(style10);
            styles1.Append(style11);
            styles1.Append(style12);
            styles1.Append(style13);

            styleDefinitionsPart1.Styles = styles1;
        }

        // Generates content of endnotesPart1.
        private void GenerateEndnotesPart1Content(EndnotesPart endnotesPart1)
        {
            Endnotes endnotes1 = new Endnotes();
            endnotes1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            endnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            endnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            endnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            endnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            endnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            endnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            endnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            endnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            Endnote endnote1 = new Endnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph87 = new Paragraph() { RsidParagraphAddition = "007252A6", RsidRunAdditionDefault = "007252A6" };

            Run run68 = new Run();
            SeparatorMark separatorMark1 = new SeparatorMark();

            run68.Append(separatorMark1);

            paragraph87.Append(run68);

            endnote1.Append(paragraph87);

            Endnote endnote2 = new Endnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph88 = new Paragraph() { RsidParagraphAddition = "007252A6", RsidRunAdditionDefault = "007252A6" };

            Run run69 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

            run69.Append(continuationSeparatorMark1);

            paragraph88.Append(run69);

            endnote2.Append(paragraph88);

            endnotes1.Append(endnote1);
            endnotes1.Append(endnote2);

            endnotesPart1.Endnotes = endnotes1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Тема Office" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Стандартная" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Стандартная" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ ゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ 明朝" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont30);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Стандартная" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint() { Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint() { Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint() { Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade() { Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade() { Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade() { Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline2 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill2);
            outline2.Append(presetDash1);

            A.Outline outline3 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill3);
            outline3.Append(presetDash2);

            A.Outline outline4 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline4.Append(solidFill4);
            outline4.Append(presetDash3);

            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);
            lineStyleList1.Append(outline4);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint() { Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 45000 };
            A.Shade shade5 = new A.Shade() { Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade() { Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade() { Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }

        // Generates content of numberingDefinitionsPart1.
        private void GenerateNumberingDefinitionsPart1Content(NumberingDefinitionsPart numberingDefinitionsPart1)
        {
            Numbering numbering1 = new Numbering();
            numbering1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            numbering1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            numbering1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            numbering1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            numbering1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            numbering1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            numbering1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            numbering1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            numbering1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 0 };
            Nsid nsid1 = new Nsid() { Val = "10D03453" };
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode1 = new TemplateCode() { Val = "9D125C1A" };

            Level level1 = new Level() { LevelIndex = 0, TemplateCode = "0419000F" };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText1 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
            Indentation indentation3 = new Indentation() { Left = "720", Hanging = "360" };

            previousParagraphProperties1.Append(indentation3);

            NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
            RunFonts runFonts79 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties1.Append(runFonts79);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);
            level1.Append(numberingSymbolRunProperties1);

            Level level2 = new Level() { LevelIndex = 1, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText2 = new LevelText() { Val = "%2." };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
            Indentation indentation4 = new Indentation() { Left = "1440", Hanging = "360" };

            previousParagraphProperties2.Append(indentation4);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);

            Level level3 = new Level() { LevelIndex = 2, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText3 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
            Indentation indentation5 = new Indentation() { Left = "2160", Hanging = "180" };

            previousParagraphProperties3.Append(indentation5);

            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);

            Level level4 = new Level() { LevelIndex = 3, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText4 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
            Indentation indentation6 = new Indentation() { Left = "2880", Hanging = "360" };

            previousParagraphProperties4.Append(indentation6);

            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);

            Level level5 = new Level() { LevelIndex = 4, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText5 = new LevelText() { Val = "%5." };
            LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
            Indentation indentation7 = new Indentation() { Left = "3600", Hanging = "360" };

            previousParagraphProperties5.Append(indentation7);

            level5.Append(startNumberingValue5);
            level5.Append(numberingFormat5);
            level5.Append(levelText5);
            level5.Append(levelJustification5);
            level5.Append(previousParagraphProperties5);

            Level level6 = new Level() { LevelIndex = 5, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText6 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
            Indentation indentation8 = new Indentation() { Left = "4320", Hanging = "180" };

            previousParagraphProperties6.Append(indentation8);

            level6.Append(startNumberingValue6);
            level6.Append(numberingFormat6);
            level6.Append(levelText6);
            level6.Append(levelJustification6);
            level6.Append(previousParagraphProperties6);

            Level level7 = new Level() { LevelIndex = 6, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText7 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
            Indentation indentation9 = new Indentation() { Left = "5040", Hanging = "360" };

            previousParagraphProperties7.Append(indentation9);

            level7.Append(startNumberingValue7);
            level7.Append(numberingFormat7);
            level7.Append(levelText7);
            level7.Append(levelJustification7);
            level7.Append(previousParagraphProperties7);

            Level level8 = new Level() { LevelIndex = 7, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText8 = new LevelText() { Val = "%8." };
            LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
            Indentation indentation10 = new Indentation() { Left = "5760", Hanging = "360" };

            previousParagraphProperties8.Append(indentation10);

            level8.Append(startNumberingValue8);
            level8.Append(numberingFormat8);
            level8.Append(levelText8);
            level8.Append(levelJustification8);
            level8.Append(previousParagraphProperties8);

            Level level9 = new Level() { LevelIndex = 8, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText9 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
            Indentation indentation11 = new Indentation() { Left = "6480", Hanging = "180" };

            previousParagraphProperties9.Append(indentation11);

            level9.Append(startNumberingValue9);
            level9.Append(numberingFormat9);
            level9.Append(levelText9);
            level9.Append(levelJustification9);
            level9.Append(previousParagraphProperties9);

            abstractNum1.Append(nsid1);
            abstractNum1.Append(multiLevelType1);
            abstractNum1.Append(templateCode1);
            abstractNum1.Append(level1);
            abstractNum1.Append(level2);
            abstractNum1.Append(level3);
            abstractNum1.Append(level4);
            abstractNum1.Append(level5);
            abstractNum1.Append(level6);
            abstractNum1.Append(level7);
            abstractNum1.Append(level8);
            abstractNum1.Append(level9);

            AbstractNum abstractNum2 = new AbstractNum() { AbstractNumberId = 1 };
            Nsid nsid2 = new Nsid() { Val = "229D2159" };
            MultiLevelType multiLevelType2 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode2 = new TemplateCode() { Val = "A52AD992" };

            Level level10 = new Level() { LevelIndex = 0, TemplateCode = "04190011" };
            StartNumberingValue startNumberingValue10 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat10 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText10 = new LevelText() { Val = "%1)" };
            LevelJustification levelJustification10 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties10 = new PreviousParagraphProperties();
            Indentation indentation12 = new Indentation() { Left = "720", Hanging = "360" };

            previousParagraphProperties10.Append(indentation12);

            NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
            RunFonts runFonts80 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties2.Append(runFonts80);

            level10.Append(startNumberingValue10);
            level10.Append(numberingFormat10);
            level10.Append(levelText10);
            level10.Append(levelJustification10);
            level10.Append(previousParagraphProperties10);
            level10.Append(numberingSymbolRunProperties2);

            Level level11 = new Level() { LevelIndex = 1, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue11 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat11 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText11 = new LevelText() { Val = "%2." };
            LevelJustification levelJustification11 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties11 = new PreviousParagraphProperties();
            Indentation indentation13 = new Indentation() { Left = "1440", Hanging = "360" };

            previousParagraphProperties11.Append(indentation13);

            level11.Append(startNumberingValue11);
            level11.Append(numberingFormat11);
            level11.Append(levelText11);
            level11.Append(levelJustification11);
            level11.Append(previousParagraphProperties11);

            Level level12 = new Level() { LevelIndex = 2, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue12 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat12 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText12 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification12 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties12 = new PreviousParagraphProperties();
            Indentation indentation14 = new Indentation() { Left = "2160", Hanging = "180" };

            previousParagraphProperties12.Append(indentation14);

            level12.Append(startNumberingValue12);
            level12.Append(numberingFormat12);
            level12.Append(levelText12);
            level12.Append(levelJustification12);
            level12.Append(previousParagraphProperties12);

            Level level13 = new Level() { LevelIndex = 3, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue13 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat13 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText13 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification13 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties13 = new PreviousParagraphProperties();
            Indentation indentation15 = new Indentation() { Left = "2880", Hanging = "360" };

            previousParagraphProperties13.Append(indentation15);

            level13.Append(startNumberingValue13);
            level13.Append(numberingFormat13);
            level13.Append(levelText13);
            level13.Append(levelJustification13);
            level13.Append(previousParagraphProperties13);

            Level level14 = new Level() { LevelIndex = 4, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue14 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat14 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText14 = new LevelText() { Val = "%5." };
            LevelJustification levelJustification14 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties14 = new PreviousParagraphProperties();
            Indentation indentation16 = new Indentation() { Left = "3600", Hanging = "360" };

            previousParagraphProperties14.Append(indentation16);

            level14.Append(startNumberingValue14);
            level14.Append(numberingFormat14);
            level14.Append(levelText14);
            level14.Append(levelJustification14);
            level14.Append(previousParagraphProperties14);

            Level level15 = new Level() { LevelIndex = 5, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue15 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat15 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText15 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification15 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties15 = new PreviousParagraphProperties();
            Indentation indentation17 = new Indentation() { Left = "4320", Hanging = "180" };

            previousParagraphProperties15.Append(indentation17);

            level15.Append(startNumberingValue15);
            level15.Append(numberingFormat15);
            level15.Append(levelText15);
            level15.Append(levelJustification15);
            level15.Append(previousParagraphProperties15);

            Level level16 = new Level() { LevelIndex = 6, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue16 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat16 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText16 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification16 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties16 = new PreviousParagraphProperties();
            Indentation indentation18 = new Indentation() { Left = "5040", Hanging = "360" };

            previousParagraphProperties16.Append(indentation18);

            level16.Append(startNumberingValue16);
            level16.Append(numberingFormat16);
            level16.Append(levelText16);
            level16.Append(levelJustification16);
            level16.Append(previousParagraphProperties16);

            Level level17 = new Level() { LevelIndex = 7, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue17 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat17 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText17 = new LevelText() { Val = "%8." };
            LevelJustification levelJustification17 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties17 = new PreviousParagraphProperties();
            Indentation indentation19 = new Indentation() { Left = "5760", Hanging = "360" };

            previousParagraphProperties17.Append(indentation19);

            level17.Append(startNumberingValue17);
            level17.Append(numberingFormat17);
            level17.Append(levelText17);
            level17.Append(levelJustification17);
            level17.Append(previousParagraphProperties17);

            Level level18 = new Level() { LevelIndex = 8, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue18 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat18 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText18 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification18 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties18 = new PreviousParagraphProperties();
            Indentation indentation20 = new Indentation() { Left = "6480", Hanging = "180" };

            previousParagraphProperties18.Append(indentation20);

            level18.Append(startNumberingValue18);
            level18.Append(numberingFormat18);
            level18.Append(levelText18);
            level18.Append(levelJustification18);
            level18.Append(previousParagraphProperties18);

            abstractNum2.Append(nsid2);
            abstractNum2.Append(multiLevelType2);
            abstractNum2.Append(templateCode2);
            abstractNum2.Append(level10);
            abstractNum2.Append(level11);
            abstractNum2.Append(level12);
            abstractNum2.Append(level13);
            abstractNum2.Append(level14);
            abstractNum2.Append(level15);
            abstractNum2.Append(level16);
            abstractNum2.Append(level17);
            abstractNum2.Append(level18);

            AbstractNum abstractNum3 = new AbstractNum() { AbstractNumberId = 2 };
            Nsid nsid3 = new Nsid() { Val = "2B7F59AF" };
            MultiLevelType multiLevelType3 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode3 = new TemplateCode() { Val = "6BE6C722" };

            Level level19 = new Level() { LevelIndex = 0, TemplateCode = "5F8AA65A" };
            StartNumberingValue startNumberingValue19 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat19 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText19 = new LevelText() { Val = "%1)" };
            LevelJustification levelJustification19 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties19 = new PreviousParagraphProperties();

            Tabs tabs3 = new Tabs();
            TabStop tabStop5 = new TabStop() { Val = TabStopValues.Number, Position = 1068 };

            tabs3.Append(tabStop5);
            Indentation indentation21 = new Indentation() { Left = "1068", Hanging = "360" };

            previousParagraphProperties19.Append(tabs3);
            previousParagraphProperties19.Append(indentation21);

            NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
            RunFonts runFonts81 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties3.Append(runFonts81);

            level19.Append(startNumberingValue19);
            level19.Append(numberingFormat19);
            level19.Append(levelText19);
            level19.Append(levelJustification19);
            level19.Append(previousParagraphProperties19);
            level19.Append(numberingSymbolRunProperties3);

            Level level20 = new Level() { LevelIndex = 1, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue20 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat20 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText20 = new LevelText() { Val = "%2." };
            LevelJustification levelJustification20 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties20 = new PreviousParagraphProperties();

            Tabs tabs4 = new Tabs();
            TabStop tabStop6 = new TabStop() { Val = TabStopValues.Number, Position = 1788 };

            tabs4.Append(tabStop6);
            Indentation indentation22 = new Indentation() { Left = "1788", Hanging = "360" };

            previousParagraphProperties20.Append(tabs4);
            previousParagraphProperties20.Append(indentation22);

            level20.Append(startNumberingValue20);
            level20.Append(numberingFormat20);
            level20.Append(levelText20);
            level20.Append(levelJustification20);
            level20.Append(previousParagraphProperties20);

            Level level21 = new Level() { LevelIndex = 2, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue21 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat21 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText21 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification21 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties21 = new PreviousParagraphProperties();

            Tabs tabs5 = new Tabs();
            TabStop tabStop7 = new TabStop() { Val = TabStopValues.Number, Position = 2508 };

            tabs5.Append(tabStop7);
            Indentation indentation23 = new Indentation() { Left = "2508", Hanging = "180" };

            previousParagraphProperties21.Append(tabs5);
            previousParagraphProperties21.Append(indentation23);

            level21.Append(startNumberingValue21);
            level21.Append(numberingFormat21);
            level21.Append(levelText21);
            level21.Append(levelJustification21);
            level21.Append(previousParagraphProperties21);

            Level level22 = new Level() { LevelIndex = 3, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue22 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat22 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText22 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification22 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties22 = new PreviousParagraphProperties();

            Tabs tabs6 = new Tabs();
            TabStop tabStop8 = new TabStop() { Val = TabStopValues.Number, Position = 3228 };

            tabs6.Append(tabStop8);
            Indentation indentation24 = new Indentation() { Left = "3228", Hanging = "360" };

            previousParagraphProperties22.Append(tabs6);
            previousParagraphProperties22.Append(indentation24);

            level22.Append(startNumberingValue22);
            level22.Append(numberingFormat22);
            level22.Append(levelText22);
            level22.Append(levelJustification22);
            level22.Append(previousParagraphProperties22);

            Level level23 = new Level() { LevelIndex = 4, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue23 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat23 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText23 = new LevelText() { Val = "%5." };
            LevelJustification levelJustification23 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties23 = new PreviousParagraphProperties();

            Tabs tabs7 = new Tabs();
            TabStop tabStop9 = new TabStop() { Val = TabStopValues.Number, Position = 3948 };

            tabs7.Append(tabStop9);
            Indentation indentation25 = new Indentation() { Left = "3948", Hanging = "360" };

            previousParagraphProperties23.Append(tabs7);
            previousParagraphProperties23.Append(indentation25);

            level23.Append(startNumberingValue23);
            level23.Append(numberingFormat23);
            level23.Append(levelText23);
            level23.Append(levelJustification23);
            level23.Append(previousParagraphProperties23);

            Level level24 = new Level() { LevelIndex = 5, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue24 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat24 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText24 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification24 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties24 = new PreviousParagraphProperties();

            Tabs tabs8 = new Tabs();
            TabStop tabStop10 = new TabStop() { Val = TabStopValues.Number, Position = 4668 };

            tabs8.Append(tabStop10);
            Indentation indentation26 = new Indentation() { Left = "4668", Hanging = "180" };

            previousParagraphProperties24.Append(tabs8);
            previousParagraphProperties24.Append(indentation26);

            level24.Append(startNumberingValue24);
            level24.Append(numberingFormat24);
            level24.Append(levelText24);
            level24.Append(levelJustification24);
            level24.Append(previousParagraphProperties24);

            Level level25 = new Level() { LevelIndex = 6, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue25 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat25 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText25 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification25 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties25 = new PreviousParagraphProperties();

            Tabs tabs9 = new Tabs();
            TabStop tabStop11 = new TabStop() { Val = TabStopValues.Number, Position = 5388 };

            tabs9.Append(tabStop11);
            Indentation indentation27 = new Indentation() { Left = "5388", Hanging = "360" };

            previousParagraphProperties25.Append(tabs9);
            previousParagraphProperties25.Append(indentation27);

            level25.Append(startNumberingValue25);
            level25.Append(numberingFormat25);
            level25.Append(levelText25);
            level25.Append(levelJustification25);
            level25.Append(previousParagraphProperties25);

            Level level26 = new Level() { LevelIndex = 7, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue26 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat26 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText26 = new LevelText() { Val = "%8." };
            LevelJustification levelJustification26 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties26 = new PreviousParagraphProperties();

            Tabs tabs10 = new Tabs();
            TabStop tabStop12 = new TabStop() { Val = TabStopValues.Number, Position = 6108 };

            tabs10.Append(tabStop12);
            Indentation indentation28 = new Indentation() { Left = "6108", Hanging = "360" };

            previousParagraphProperties26.Append(tabs10);
            previousParagraphProperties26.Append(indentation28);

            level26.Append(startNumberingValue26);
            level26.Append(numberingFormat26);
            level26.Append(levelText26);
            level26.Append(levelJustification26);
            level26.Append(previousParagraphProperties26);

            Level level27 = new Level() { LevelIndex = 8, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue27 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat27 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText27 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification27 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties27 = new PreviousParagraphProperties();

            Tabs tabs11 = new Tabs();
            TabStop tabStop13 = new TabStop() { Val = TabStopValues.Number, Position = 6828 };

            tabs11.Append(tabStop13);
            Indentation indentation29 = new Indentation() { Left = "6828", Hanging = "180" };

            previousParagraphProperties27.Append(tabs11);
            previousParagraphProperties27.Append(indentation29);

            level27.Append(startNumberingValue27);
            level27.Append(numberingFormat27);
            level27.Append(levelText27);
            level27.Append(levelJustification27);
            level27.Append(previousParagraphProperties27);

            abstractNum3.Append(nsid3);
            abstractNum3.Append(multiLevelType3);
            abstractNum3.Append(templateCode3);
            abstractNum3.Append(level19);
            abstractNum3.Append(level20);
            abstractNum3.Append(level21);
            abstractNum3.Append(level22);
            abstractNum3.Append(level23);
            abstractNum3.Append(level24);
            abstractNum3.Append(level25);
            abstractNum3.Append(level26);
            abstractNum3.Append(level27);

            AbstractNum abstractNum4 = new AbstractNum() { AbstractNumberId = 3 };
            Nsid nsid4 = new Nsid() { Val = "4C4F4639" };
            MultiLevelType multiLevelType4 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode4 = new TemplateCode() { Val = "93FCBAA8" };

            Level level28 = new Level() { LevelIndex = 0, TemplateCode = "04190001" };
            StartNumberingValue startNumberingValue28 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat28 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText28 = new LevelText() { Val = "·" };
            LevelJustification levelJustification28 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties28 = new PreviousParagraphProperties();

            Tabs tabs12 = new Tabs();
            TabStop tabStop14 = new TabStop() { Val = TabStopValues.Number, Position = 720 };

            tabs12.Append(tabStop14);
            Indentation indentation30 = new Indentation() { Left = "720", Hanging = "360" };

            previousParagraphProperties28.Append(tabs12);
            previousParagraphProperties28.Append(indentation30);

            NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
            RunFonts runFonts82 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

            numberingSymbolRunProperties4.Append(runFonts82);

            level28.Append(startNumberingValue28);
            level28.Append(numberingFormat28);
            level28.Append(levelText28);
            level28.Append(levelJustification28);
            level28.Append(previousParagraphProperties28);
            level28.Append(numberingSymbolRunProperties4);

            Level level29 = new Level() { LevelIndex = 1, TemplateCode = "04190003", Tentative = true };
            StartNumberingValue startNumberingValue29 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat29 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText29 = new LevelText() { Val = "o" };
            LevelJustification levelJustification29 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties29 = new PreviousParagraphProperties();

            Tabs tabs13 = new Tabs();
            TabStop tabStop15 = new TabStop() { Val = TabStopValues.Number, Position = 1440 };

            tabs13.Append(tabStop15);
            Indentation indentation31 = new Indentation() { Left = "1440", Hanging = "360" };

            previousParagraphProperties29.Append(tabs13);
            previousParagraphProperties29.Append(indentation31);

            NumberingSymbolRunProperties numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();
            RunFonts runFonts83 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };

            numberingSymbolRunProperties5.Append(runFonts83);

            level29.Append(startNumberingValue29);
            level29.Append(numberingFormat29);
            level29.Append(levelText29);
            level29.Append(levelJustification29);
            level29.Append(previousParagraphProperties29);
            level29.Append(numberingSymbolRunProperties5);

            Level level30 = new Level() { LevelIndex = 2, TemplateCode = "04190005", Tentative = true };
            StartNumberingValue startNumberingValue30 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat30 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText30 = new LevelText() { Val = "§" };
            LevelJustification levelJustification30 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties30 = new PreviousParagraphProperties();

            Tabs tabs14 = new Tabs();
            TabStop tabStop16 = new TabStop() { Val = TabStopValues.Number, Position = 2160 };

            tabs14.Append(tabStop16);
            Indentation indentation32 = new Indentation() { Left = "2160", Hanging = "360" };

            previousParagraphProperties30.Append(tabs14);
            previousParagraphProperties30.Append(indentation32);

            NumberingSymbolRunProperties numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();
            RunFonts runFonts84 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties6.Append(runFonts84);

            level30.Append(startNumberingValue30);
            level30.Append(numberingFormat30);
            level30.Append(levelText30);
            level30.Append(levelJustification30);
            level30.Append(previousParagraphProperties30);
            level30.Append(numberingSymbolRunProperties6);

            Level level31 = new Level() { LevelIndex = 3, TemplateCode = "04190001", Tentative = true };
            StartNumberingValue startNumberingValue31 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat31 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText31 = new LevelText() { Val = "·" };
            LevelJustification levelJustification31 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties31 = new PreviousParagraphProperties();

            Tabs tabs15 = new Tabs();
            TabStop tabStop17 = new TabStop() { Val = TabStopValues.Number, Position = 2880 };

            tabs15.Append(tabStop17);
            Indentation indentation33 = new Indentation() { Left = "2880", Hanging = "360" };

            previousParagraphProperties31.Append(tabs15);
            previousParagraphProperties31.Append(indentation33);

            NumberingSymbolRunProperties numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();
            RunFonts runFonts85 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

            numberingSymbolRunProperties7.Append(runFonts85);

            level31.Append(startNumberingValue31);
            level31.Append(numberingFormat31);
            level31.Append(levelText31);
            level31.Append(levelJustification31);
            level31.Append(previousParagraphProperties31);
            level31.Append(numberingSymbolRunProperties7);

            Level level32 = new Level() { LevelIndex = 4, TemplateCode = "04190003", Tentative = true };
            StartNumberingValue startNumberingValue32 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat32 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText32 = new LevelText() { Val = "o" };
            LevelJustification levelJustification32 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties32 = new PreviousParagraphProperties();

            Tabs tabs16 = new Tabs();
            TabStop tabStop18 = new TabStop() { Val = TabStopValues.Number, Position = 3600 };

            tabs16.Append(tabStop18);
            Indentation indentation34 = new Indentation() { Left = "3600", Hanging = "360" };

            previousParagraphProperties32.Append(tabs16);
            previousParagraphProperties32.Append(indentation34);

            NumberingSymbolRunProperties numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();
            RunFonts runFonts86 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };

            numberingSymbolRunProperties8.Append(runFonts86);

            level32.Append(startNumberingValue32);
            level32.Append(numberingFormat32);
            level32.Append(levelText32);
            level32.Append(levelJustification32);
            level32.Append(previousParagraphProperties32);
            level32.Append(numberingSymbolRunProperties8);

            Level level33 = new Level() { LevelIndex = 5, TemplateCode = "04190005", Tentative = true };
            StartNumberingValue startNumberingValue33 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat33 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText33 = new LevelText() { Val = "§" };
            LevelJustification levelJustification33 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties33 = new PreviousParagraphProperties();

            Tabs tabs17 = new Tabs();
            TabStop tabStop19 = new TabStop() { Val = TabStopValues.Number, Position = 4320 };

            tabs17.Append(tabStop19);
            Indentation indentation35 = new Indentation() { Left = "4320", Hanging = "360" };

            previousParagraphProperties33.Append(tabs17);
            previousParagraphProperties33.Append(indentation35);

            NumberingSymbolRunProperties numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();
            RunFonts runFonts87 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties9.Append(runFonts87);

            level33.Append(startNumberingValue33);
            level33.Append(numberingFormat33);
            level33.Append(levelText33);
            level33.Append(levelJustification33);
            level33.Append(previousParagraphProperties33);
            level33.Append(numberingSymbolRunProperties9);

            Level level34 = new Level() { LevelIndex = 6, TemplateCode = "04190001", Tentative = true };
            StartNumberingValue startNumberingValue34 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat34 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText34 = new LevelText() { Val = "·" };
            LevelJustification levelJustification34 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties34 = new PreviousParagraphProperties();

            Tabs tabs18 = new Tabs();
            TabStop tabStop20 = new TabStop() { Val = TabStopValues.Number, Position = 5040 };

            tabs18.Append(tabStop20);
            Indentation indentation36 = new Indentation() { Left = "5040", Hanging = "360" };

            previousParagraphProperties34.Append(tabs18);
            previousParagraphProperties34.Append(indentation36);

            NumberingSymbolRunProperties numberingSymbolRunProperties10 = new NumberingSymbolRunProperties();
            RunFonts runFonts88 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

            numberingSymbolRunProperties10.Append(runFonts88);

            level34.Append(startNumberingValue34);
            level34.Append(numberingFormat34);
            level34.Append(levelText34);
            level34.Append(levelJustification34);
            level34.Append(previousParagraphProperties34);
            level34.Append(numberingSymbolRunProperties10);

            Level level35 = new Level() { LevelIndex = 7, TemplateCode = "04190003", Tentative = true };
            StartNumberingValue startNumberingValue35 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat35 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText35 = new LevelText() { Val = "o" };
            LevelJustification levelJustification35 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties35 = new PreviousParagraphProperties();

            Tabs tabs19 = new Tabs();
            TabStop tabStop21 = new TabStop() { Val = TabStopValues.Number, Position = 5760 };

            tabs19.Append(tabStop21);
            Indentation indentation37 = new Indentation() { Left = "5760", Hanging = "360" };

            previousParagraphProperties35.Append(tabs19);
            previousParagraphProperties35.Append(indentation37);

            NumberingSymbolRunProperties numberingSymbolRunProperties11 = new NumberingSymbolRunProperties();
            RunFonts runFonts89 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };

            numberingSymbolRunProperties11.Append(runFonts89);

            level35.Append(startNumberingValue35);
            level35.Append(numberingFormat35);
            level35.Append(levelText35);
            level35.Append(levelJustification35);
            level35.Append(previousParagraphProperties35);
            level35.Append(numberingSymbolRunProperties11);

            Level level36 = new Level() { LevelIndex = 8, TemplateCode = "04190005", Tentative = true };
            StartNumberingValue startNumberingValue36 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat36 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText36 = new LevelText() { Val = "§" };
            LevelJustification levelJustification36 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties36 = new PreviousParagraphProperties();

            Tabs tabs20 = new Tabs();
            TabStop tabStop22 = new TabStop() { Val = TabStopValues.Number, Position = 6480 };

            tabs20.Append(tabStop22);
            Indentation indentation38 = new Indentation() { Left = "6480", Hanging = "360" };

            previousParagraphProperties36.Append(tabs20);
            previousParagraphProperties36.Append(indentation38);

            NumberingSymbolRunProperties numberingSymbolRunProperties12 = new NumberingSymbolRunProperties();
            RunFonts runFonts90 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties12.Append(runFonts90);

            level36.Append(startNumberingValue36);
            level36.Append(numberingFormat36);
            level36.Append(levelText36);
            level36.Append(levelJustification36);
            level36.Append(previousParagraphProperties36);
            level36.Append(numberingSymbolRunProperties12);

            abstractNum4.Append(nsid4);
            abstractNum4.Append(multiLevelType4);
            abstractNum4.Append(templateCode4);
            abstractNum4.Append(level28);
            abstractNum4.Append(level29);
            abstractNum4.Append(level30);
            abstractNum4.Append(level31);
            abstractNum4.Append(level32);
            abstractNum4.Append(level33);
            abstractNum4.Append(level34);
            abstractNum4.Append(level35);
            abstractNum4.Append(level36);

            AbstractNum abstractNum5 = new AbstractNum() { AbstractNumberId = 4 };
            Nsid nsid5 = new Nsid() { Val = "5599259D" };
            MultiLevelType multiLevelType5 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode5 = new TemplateCode() { Val = "207A637E" };

            Level level37 = new Level() { LevelIndex = 0, TemplateCode = "11F2B736" };
            StartNumberingValue startNumberingValue37 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat37 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText37 = new LevelText() { Val = "%1)" };
            LevelJustification levelJustification37 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties37 = new PreviousParagraphProperties();
            Indentation indentation39 = new Indentation() { Left = "1211", Hanging = "360" };

            previousParagraphProperties37.Append(indentation39);

            NumberingSymbolRunProperties numberingSymbolRunProperties13 = new NumberingSymbolRunProperties();
            RunFonts runFonts91 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties13.Append(runFonts91);

            level37.Append(startNumberingValue37);
            level37.Append(numberingFormat37);
            level37.Append(levelText37);
            level37.Append(levelJustification37);
            level37.Append(previousParagraphProperties37);
            level37.Append(numberingSymbolRunProperties13);

            Level level38 = new Level() { LevelIndex = 1, TemplateCode = "04190019" };
            StartNumberingValue startNumberingValue38 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat38 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText38 = new LevelText() { Val = "%2." };
            LevelJustification levelJustification38 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties38 = new PreviousParagraphProperties();
            Indentation indentation40 = new Indentation() { Left = "1931", Hanging = "360" };

            previousParagraphProperties38.Append(indentation40);

            level38.Append(startNumberingValue38);
            level38.Append(numberingFormat38);
            level38.Append(levelText38);
            level38.Append(levelJustification38);
            level38.Append(previousParagraphProperties38);

            Level level39 = new Level() { LevelIndex = 2, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue39 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat39 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText39 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification39 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties39 = new PreviousParagraphProperties();
            Indentation indentation41 = new Indentation() { Left = "2651", Hanging = "180" };

            previousParagraphProperties39.Append(indentation41);

            level39.Append(startNumberingValue39);
            level39.Append(numberingFormat39);
            level39.Append(levelText39);
            level39.Append(levelJustification39);
            level39.Append(previousParagraphProperties39);

            Level level40 = new Level() { LevelIndex = 3, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue40 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat40 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText40 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification40 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties40 = new PreviousParagraphProperties();
            Indentation indentation42 = new Indentation() { Left = "3371", Hanging = "360" };

            previousParagraphProperties40.Append(indentation42);

            level40.Append(startNumberingValue40);
            level40.Append(numberingFormat40);
            level40.Append(levelText40);
            level40.Append(levelJustification40);
            level40.Append(previousParagraphProperties40);

            Level level41 = new Level() { LevelIndex = 4, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue41 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat41 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText41 = new LevelText() { Val = "%5." };
            LevelJustification levelJustification41 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties41 = new PreviousParagraphProperties();
            Indentation indentation43 = new Indentation() { Left = "4091", Hanging = "360" };

            previousParagraphProperties41.Append(indentation43);

            level41.Append(startNumberingValue41);
            level41.Append(numberingFormat41);
            level41.Append(levelText41);
            level41.Append(levelJustification41);
            level41.Append(previousParagraphProperties41);

            Level level42 = new Level() { LevelIndex = 5, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue42 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat42 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText42 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification42 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties42 = new PreviousParagraphProperties();
            Indentation indentation44 = new Indentation() { Left = "4811", Hanging = "180" };

            previousParagraphProperties42.Append(indentation44);

            level42.Append(startNumberingValue42);
            level42.Append(numberingFormat42);
            level42.Append(levelText42);
            level42.Append(levelJustification42);
            level42.Append(previousParagraphProperties42);

            Level level43 = new Level() { LevelIndex = 6, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue43 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat43 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText43 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification43 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties43 = new PreviousParagraphProperties();
            Indentation indentation45 = new Indentation() { Left = "5531", Hanging = "360" };

            previousParagraphProperties43.Append(indentation45);

            level43.Append(startNumberingValue43);
            level43.Append(numberingFormat43);
            level43.Append(levelText43);
            level43.Append(levelJustification43);
            level43.Append(previousParagraphProperties43);

            Level level44 = new Level() { LevelIndex = 7, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue44 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat44 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText44 = new LevelText() { Val = "%8." };
            LevelJustification levelJustification44 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties44 = new PreviousParagraphProperties();
            Indentation indentation46 = new Indentation() { Left = "6251", Hanging = "360" };

            previousParagraphProperties44.Append(indentation46);

            level44.Append(startNumberingValue44);
            level44.Append(numberingFormat44);
            level44.Append(levelText44);
            level44.Append(levelJustification44);
            level44.Append(previousParagraphProperties44);

            Level level45 = new Level() { LevelIndex = 8, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue45 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat45 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText45 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification45 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties45 = new PreviousParagraphProperties();
            Indentation indentation47 = new Indentation() { Left = "6971", Hanging = "180" };

            previousParagraphProperties45.Append(indentation47);

            level45.Append(startNumberingValue45);
            level45.Append(numberingFormat45);
            level45.Append(levelText45);
            level45.Append(levelJustification45);
            level45.Append(previousParagraphProperties45);

            abstractNum5.Append(nsid5);
            abstractNum5.Append(multiLevelType5);
            abstractNum5.Append(templateCode5);
            abstractNum5.Append(level37);
            abstractNum5.Append(level38);
            abstractNum5.Append(level39);
            abstractNum5.Append(level40);
            abstractNum5.Append(level41);
            abstractNum5.Append(level42);
            abstractNum5.Append(level43);
            abstractNum5.Append(level44);
            abstractNum5.Append(level45);

            AbstractNum abstractNum6 = new AbstractNum() { AbstractNumberId = 5 };
            Nsid nsid6 = new Nsid() { Val = "7F872B71" };
            MultiLevelType multiLevelType6 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode6 = new TemplateCode() { Val = "8BE8B986" };

            Level level46 = new Level() { LevelIndex = 0, TemplateCode = "04190001" };
            StartNumberingValue startNumberingValue46 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat46 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText46 = new LevelText() { Val = "·" };
            LevelJustification levelJustification46 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties46 = new PreviousParagraphProperties();

            Tabs tabs21 = new Tabs();
            TabStop tabStop23 = new TabStop() { Val = TabStopValues.Number, Position = 1428 };

            tabs21.Append(tabStop23);
            Indentation indentation48 = new Indentation() { Left = "1428", Hanging = "360" };

            previousParagraphProperties46.Append(tabs21);
            previousParagraphProperties46.Append(indentation48);

            NumberingSymbolRunProperties numberingSymbolRunProperties14 = new NumberingSymbolRunProperties();
            RunFonts runFonts92 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

            numberingSymbolRunProperties14.Append(runFonts92);

            level46.Append(startNumberingValue46);
            level46.Append(numberingFormat46);
            level46.Append(levelText46);
            level46.Append(levelJustification46);
            level46.Append(previousParagraphProperties46);
            level46.Append(numberingSymbolRunProperties14);

            Level level47 = new Level() { LevelIndex = 1, TemplateCode = "04190003", Tentative = true };
            StartNumberingValue startNumberingValue47 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat47 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText47 = new LevelText() { Val = "o" };
            LevelJustification levelJustification47 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties47 = new PreviousParagraphProperties();

            Tabs tabs22 = new Tabs();
            TabStop tabStop24 = new TabStop() { Val = TabStopValues.Number, Position = 2148 };

            tabs22.Append(tabStop24);
            Indentation indentation49 = new Indentation() { Left = "2148", Hanging = "360" };

            previousParagraphProperties47.Append(tabs22);
            previousParagraphProperties47.Append(indentation49);

            NumberingSymbolRunProperties numberingSymbolRunProperties15 = new NumberingSymbolRunProperties();
            RunFonts runFonts93 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };

            numberingSymbolRunProperties15.Append(runFonts93);

            level47.Append(startNumberingValue47);
            level47.Append(numberingFormat47);
            level47.Append(levelText47);
            level47.Append(levelJustification47);
            level47.Append(previousParagraphProperties47);
            level47.Append(numberingSymbolRunProperties15);

            Level level48 = new Level() { LevelIndex = 2, TemplateCode = "04190005", Tentative = true };
            StartNumberingValue startNumberingValue48 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat48 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText48 = new LevelText() { Val = "§" };
            LevelJustification levelJustification48 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties48 = new PreviousParagraphProperties();

            Tabs tabs23 = new Tabs();
            TabStop tabStop25 = new TabStop() { Val = TabStopValues.Number, Position = 2868 };

            tabs23.Append(tabStop25);
            Indentation indentation50 = new Indentation() { Left = "2868", Hanging = "360" };

            previousParagraphProperties48.Append(tabs23);
            previousParagraphProperties48.Append(indentation50);

            NumberingSymbolRunProperties numberingSymbolRunProperties16 = new NumberingSymbolRunProperties();
            RunFonts runFonts94 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties16.Append(runFonts94);

            level48.Append(startNumberingValue48);
            level48.Append(numberingFormat48);
            level48.Append(levelText48);
            level48.Append(levelJustification48);
            level48.Append(previousParagraphProperties48);
            level48.Append(numberingSymbolRunProperties16);

            Level level49 = new Level() { LevelIndex = 3, TemplateCode = "04190001", Tentative = true };
            StartNumberingValue startNumberingValue49 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat49 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText49 = new LevelText() { Val = "·" };
            LevelJustification levelJustification49 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties49 = new PreviousParagraphProperties();

            Tabs tabs24 = new Tabs();
            TabStop tabStop26 = new TabStop() { Val = TabStopValues.Number, Position = 3588 };

            tabs24.Append(tabStop26);
            Indentation indentation51 = new Indentation() { Left = "3588", Hanging = "360" };

            previousParagraphProperties49.Append(tabs24);
            previousParagraphProperties49.Append(indentation51);

            NumberingSymbolRunProperties numberingSymbolRunProperties17 = new NumberingSymbolRunProperties();
            RunFonts runFonts95 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

            numberingSymbolRunProperties17.Append(runFonts95);

            level49.Append(startNumberingValue49);
            level49.Append(numberingFormat49);
            level49.Append(levelText49);
            level49.Append(levelJustification49);
            level49.Append(previousParagraphProperties49);
            level49.Append(numberingSymbolRunProperties17);

            Level level50 = new Level() { LevelIndex = 4, TemplateCode = "04190003", Tentative = true };
            StartNumberingValue startNumberingValue50 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat50 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText50 = new LevelText() { Val = "o" };
            LevelJustification levelJustification50 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties50 = new PreviousParagraphProperties();

            Tabs tabs25 = new Tabs();
            TabStop tabStop27 = new TabStop() { Val = TabStopValues.Number, Position = 4308 };

            tabs25.Append(tabStop27);
            Indentation indentation52 = new Indentation() { Left = "4308", Hanging = "360" };

            previousParagraphProperties50.Append(tabs25);
            previousParagraphProperties50.Append(indentation52);

            NumberingSymbolRunProperties numberingSymbolRunProperties18 = new NumberingSymbolRunProperties();
            RunFonts runFonts96 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };

            numberingSymbolRunProperties18.Append(runFonts96);

            level50.Append(startNumberingValue50);
            level50.Append(numberingFormat50);
            level50.Append(levelText50);
            level50.Append(levelJustification50);
            level50.Append(previousParagraphProperties50);
            level50.Append(numberingSymbolRunProperties18);

            Level level51 = new Level() { LevelIndex = 5, TemplateCode = "04190005", Tentative = true };
            StartNumberingValue startNumberingValue51 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat51 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText51 = new LevelText() { Val = "§" };
            LevelJustification levelJustification51 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties51 = new PreviousParagraphProperties();

            Tabs tabs26 = new Tabs();
            TabStop tabStop28 = new TabStop() { Val = TabStopValues.Number, Position = 5028 };

            tabs26.Append(tabStop28);
            Indentation indentation53 = new Indentation() { Left = "5028", Hanging = "360" };

            previousParagraphProperties51.Append(tabs26);
            previousParagraphProperties51.Append(indentation53);

            NumberingSymbolRunProperties numberingSymbolRunProperties19 = new NumberingSymbolRunProperties();
            RunFonts runFonts97 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties19.Append(runFonts97);

            level51.Append(startNumberingValue51);
            level51.Append(numberingFormat51);
            level51.Append(levelText51);
            level51.Append(levelJustification51);
            level51.Append(previousParagraphProperties51);
            level51.Append(numberingSymbolRunProperties19);

            Level level52 = new Level() { LevelIndex = 6, TemplateCode = "04190001", Tentative = true };
            StartNumberingValue startNumberingValue52 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat52 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText52 = new LevelText() { Val = "·" };
            LevelJustification levelJustification52 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties52 = new PreviousParagraphProperties();

            Tabs tabs27 = new Tabs();
            TabStop tabStop29 = new TabStop() { Val = TabStopValues.Number, Position = 5748 };

            tabs27.Append(tabStop29);
            Indentation indentation54 = new Indentation() { Left = "5748", Hanging = "360" };

            previousParagraphProperties52.Append(tabs27);
            previousParagraphProperties52.Append(indentation54);

            NumberingSymbolRunProperties numberingSymbolRunProperties20 = new NumberingSymbolRunProperties();
            RunFonts runFonts98 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

            numberingSymbolRunProperties20.Append(runFonts98);

            level52.Append(startNumberingValue52);
            level52.Append(numberingFormat52);
            level52.Append(levelText52);
            level52.Append(levelJustification52);
            level52.Append(previousParagraphProperties52);
            level52.Append(numberingSymbolRunProperties20);

            Level level53 = new Level() { LevelIndex = 7, TemplateCode = "04190003", Tentative = true };
            StartNumberingValue startNumberingValue53 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat53 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText53 = new LevelText() { Val = "o" };
            LevelJustification levelJustification53 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties53 = new PreviousParagraphProperties();

            Tabs tabs28 = new Tabs();
            TabStop tabStop30 = new TabStop() { Val = TabStopValues.Number, Position = 6468 };

            tabs28.Append(tabStop30);
            Indentation indentation55 = new Indentation() { Left = "6468", Hanging = "360" };

            previousParagraphProperties53.Append(tabs28);
            previousParagraphProperties53.Append(indentation55);

            NumberingSymbolRunProperties numberingSymbolRunProperties21 = new NumberingSymbolRunProperties();
            RunFonts runFonts99 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };

            numberingSymbolRunProperties21.Append(runFonts99);

            level53.Append(startNumberingValue53);
            level53.Append(numberingFormat53);
            level53.Append(levelText53);
            level53.Append(levelJustification53);
            level53.Append(previousParagraphProperties53);
            level53.Append(numberingSymbolRunProperties21);

            Level level54 = new Level() { LevelIndex = 8, TemplateCode = "04190005", Tentative = true };
            StartNumberingValue startNumberingValue54 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat54 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText54 = new LevelText() { Val = "§" };
            LevelJustification levelJustification54 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties54 = new PreviousParagraphProperties();

            Tabs tabs29 = new Tabs();
            TabStop tabStop31 = new TabStop() { Val = TabStopValues.Number, Position = 7188 };

            tabs29.Append(tabStop31);
            Indentation indentation56 = new Indentation() { Left = "7188", Hanging = "360" };

            previousParagraphProperties54.Append(tabs29);
            previousParagraphProperties54.Append(indentation56);

            NumberingSymbolRunProperties numberingSymbolRunProperties22 = new NumberingSymbolRunProperties();
            RunFonts runFonts100 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties22.Append(runFonts100);

            level54.Append(startNumberingValue54);
            level54.Append(numberingFormat54);
            level54.Append(levelText54);
            level54.Append(levelJustification54);
            level54.Append(previousParagraphProperties54);
            level54.Append(numberingSymbolRunProperties22);

            abstractNum6.Append(nsid6);
            abstractNum6.Append(multiLevelType6);
            abstractNum6.Append(templateCode6);
            abstractNum6.Append(level46);
            abstractNum6.Append(level47);
            abstractNum6.Append(level48);
            abstractNum6.Append(level49);
            abstractNum6.Append(level50);
            abstractNum6.Append(level51);
            abstractNum6.Append(level52);
            abstractNum6.Append(level53);
            abstractNum6.Append(level54);

            NumberingInstance numberingInstance1 = new NumberingInstance() { NumberID = 1 };
            AbstractNumId abstractNumId1 = new AbstractNumId() { Val = 2 };

            numberingInstance1.Append(abstractNumId1);

            NumberingInstance numberingInstance2 = new NumberingInstance() { NumberID = 2 };
            AbstractNumId abstractNumId2 = new AbstractNumId() { Val = 5 };

            numberingInstance2.Append(abstractNumId2);

            NumberingInstance numberingInstance3 = new NumberingInstance() { NumberID = 3 };
            AbstractNumId abstractNumId3 = new AbstractNumId() { Val = 3 };

            numberingInstance3.Append(abstractNumId3);

            NumberingInstance numberingInstance4 = new NumberingInstance() { NumberID = 4 };
            AbstractNumId abstractNumId4 = new AbstractNumId() { Val = 0 };

            numberingInstance4.Append(abstractNumId4);

            NumberingInstance numberingInstance5 = new NumberingInstance() { NumberID = 5 };
            AbstractNumId abstractNumId5 = new AbstractNumId() { Val = 4 };

            numberingInstance5.Append(abstractNumId5);

            NumberingInstance numberingInstance6 = new NumberingInstance() { NumberID = 6 };
            AbstractNumId abstractNumId6 = new AbstractNumId() { Val = 1 };

            numberingInstance6.Append(abstractNumId6);

            numbering1.Append(abstractNum1);
            numbering1.Append(abstractNum2);
            numbering1.Append(abstractNum3);
            numbering1.Append(abstractNum4);
            numbering1.Append(abstractNum5);
            numbering1.Append(abstractNum6);
            numbering1.Append(numberingInstance1);
            numbering1.Append(numberingInstance2);
            numbering1.Append(numberingInstance3);
            numbering1.Append(numberingInstance4);
            numbering1.Append(numberingInstance5);
            numbering1.Append(numberingInstance6);

            numberingDefinitionsPart1.Numbering = numbering1;
        }

        // Generates content of customXmlPart1.
        private void GenerateCustomXmlPart1Content(CustomXmlPart customXmlPart1)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(customXmlPart1.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<b:Sources SelectedStyle=\"\\APA.XSL\" StyleName=\"APA\" xmlns:b=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\" xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\"></b:Sources>\r\n");
            writer.Flush();
            writer.Close();
        }

        // Generates content of customXmlPropertiesPart1.
        private void GenerateCustomXmlPropertiesPart1Content(CustomXmlPropertiesPart customXmlPropertiesPart1)
        {
            Ds.DataStoreItem dataStoreItem1 = new Ds.DataStoreItem() { ItemId = "{9B602245-1C13-4C17-95B8-2E916E2BE159}" };
            dataStoreItem1.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

            Ds.SchemaReferences schemaReferences1 = new Ds.SchemaReferences();
            Ds.SchemaReference schemaReference1 = new Ds.SchemaReference() { Uri = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography" };

            schemaReferences1.Append(schemaReference1);

            dataStoreItem1.Append(schemaReferences1);

            customXmlPropertiesPart1.DataStoreItem = dataStoreItem1;
        }

        // Generates content of footnotesPart1.
        private void GenerateFootnotesPart1Content(FootnotesPart footnotesPart1)
        {
            Footnotes footnotes1 = new Footnotes();
            footnotes1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            Footnote footnote1 = new Footnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph89 = new Paragraph() { RsidParagraphAddition = "007252A6", RsidRunAdditionDefault = "007252A6" };

            Run run70 = new Run();
            SeparatorMark separatorMark2 = new SeparatorMark();

            run70.Append(separatorMark2);

            paragraph89.Append(run70);

            footnote1.Append(paragraph89);

            Footnote footnote2 = new Footnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph90 = new Paragraph() { RsidParagraphAddition = "007252A6", RsidRunAdditionDefault = "007252A6" };

            Run run71 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark2 = new ContinuationSeparatorMark();

            run71.Append(continuationSeparatorMark2);

            paragraph90.Append(run71);

            footnote2.Append(paragraph90);

            footnotes1.Append(footnote1);
            footnotes1.Append(footnote2);

            footnotesPart1.Footnotes = footnotes1;
        }

        // Generates content of fontTablePart1.
        private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts();
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            Font font1 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C0007841", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "Symbol" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "05050102010706020507" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "02" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "00000000", UnicodeSignature1 = "10000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "80000000", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "Courier New" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "02070309020205020404" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Modern };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Fixed };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C0007843", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "Wingdings" };
            Panose1Number panose1Number4 = new Panose1Number() { Val = "05000000000000000000" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "02" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Auto };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "00000000", UnicodeSignature1 = "10000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "80000000", CodePageSignature1 = "00000000" };

            font4.Append(panose1Number4);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            Font font5 = new Font() { Name = "HeliosCond" };
            Panose1Number panose1Number5 = new Panose1Number() { Val = "00000000000000000000" };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily5 = new FontFamily() { Val = FontFamilyValues.Decorative };
            NotTrueType notTrueType1 = new NotTrueType();
            Pitch pitch5 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature5 = new FontSignature() { UnicodeSignature0 = "00000203", UnicodeSignature1 = "00000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "00000005", CodePageSignature1 = "00000000" };

            font5.Append(panose1Number5);
            font5.Append(fontCharSet5);
            font5.Append(fontFamily5);
            font5.Append(notTrueType1);
            font5.Append(pitch5);
            font5.Append(fontSignature5);

            Font font6 = new Font() { Name = "Tahoma" };
            Panose1Number panose1Number6 = new Panose1Number() { Val = "020B0604030504040204" };
            FontCharSet fontCharSet6 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily6 = new FontFamily() { Val = FontFamilyValues.Swiss };
            NotTrueType notTrueType2 = new NotTrueType();
            Pitch pitch6 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature6 = new FontSignature() { UnicodeSignature0 = "00000003", UnicodeSignature1 = "00000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "00000001", CodePageSignature1 = "00000000" };

            font6.Append(panose1Number6);
            font6.Append(fontCharSet6);
            font6.Append(fontFamily6);
            font6.Append(notTrueType2);
            font6.Append(pitch6);
            font6.Append(fontSignature6);

            Font font7 = new Font() { Name = "HeliosCond-Bold" };
            Panose1Number panose1Number7 = new Panose1Number() { Val = "00000000000000000000" };
            FontCharSet fontCharSet7 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily7 = new FontFamily() { Val = FontFamilyValues.Auto };
            NotTrueType notTrueType3 = new NotTrueType();
            Pitch pitch7 = new Pitch() { Val = FontPitchValues.Default };
            FontSignature fontSignature7 = new FontSignature() { UnicodeSignature0 = "00000201", UnicodeSignature1 = "00000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "00000004", CodePageSignature1 = "00000000" };

            font7.Append(panose1Number7);
            font7.Append(fontCharSet7);
            font7.Append(fontFamily7);
            font7.Append(notTrueType3);
            font7.Append(pitch7);
            font7.Append(fontSignature7);

            Font font8 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number8 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet8 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily8 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch8 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature8 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "4000ACFF", UnicodeSignature2 = "00000001", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font8.Append(panose1Number8);
            font8.Append(fontCharSet8);
            font8.Append(fontFamily8);
            font8.Append(pitch8);
            font8.Append(fontSignature8);

            Font font9 = new Font() { Name = "Courier New CYR" };
            Panose1Number panose1Number9 = new Panose1Number() { Val = "02070309020205020404" };
            FontCharSet fontCharSet9 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily9 = new FontFamily() { Val = FontFamilyValues.Modern };
            Pitch pitch9 = new Pitch() { Val = FontPitchValues.Fixed };
            FontSignature fontSignature9 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C0007843", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font9.Append(panose1Number9);
            font9.Append(fontCharSet9);
            font9.Append(fontFamily9);
            font9.Append(pitch9);
            font9.Append(fontSignature9);

            Font font10 = new Font() { Name = "Cambria" };
            Panose1Number panose1Number10 = new Panose1Number() { Val = "02040503050406030204" };
            FontCharSet fontCharSet10 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily10 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch10 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature10 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "400004FF", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font10.Append(panose1Number10);
            font10.Append(fontCharSet10);
            font10.Append(fontFamily10);
            font10.Append(pitch10);
            font10.Append(fontSignature10);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);
            fonts1.Append(font7);
            fonts1.Append(font8);
            fonts1.Append(font9);
            fonts1.Append(font10);

            fontTablePart1.Fonts = fonts1;
        }

        // Generates content of webSettingsPart1.
        private void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = new WebSettings();
            webSettings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            webSettings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            Divs divs1 = new Divs();

            Div div1 = new Div() { Id = "192423841" };
            BodyDiv bodyDiv1 = new BodyDiv() { Val = true };
            LeftMarginDiv leftMarginDiv1 = new LeftMarginDiv() { Val = "0" };
            RightMarginDiv rightMarginDiv1 = new RightMarginDiv() { Val = "0" };
            TopMarginDiv topMarginDiv1 = new TopMarginDiv() { Val = "0" };
            BottomMarginDiv bottomMarginDiv1 = new BottomMarginDiv() { Val = "0" };

            DivBorder divBorder1 = new DivBorder();
            TopBorder topBorder6 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder7 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder1.Append(topBorder6);
            divBorder1.Append(leftBorder2);
            divBorder1.Append(bottomBorder7);
            divBorder1.Append(rightBorder2);

            div1.Append(bodyDiv1);
            div1.Append(leftMarginDiv1);
            div1.Append(rightMarginDiv1);
            div1.Append(topMarginDiv1);
            div1.Append(bottomMarginDiv1);
            div1.Append(divBorder1);

            Div div2 = new Div() { Id = "1366713195" };
            BodyDiv bodyDiv2 = new BodyDiv() { Val = true };
            LeftMarginDiv leftMarginDiv2 = new LeftMarginDiv() { Val = "0" };
            RightMarginDiv rightMarginDiv2 = new RightMarginDiv() { Val = "0" };
            TopMarginDiv topMarginDiv2 = new TopMarginDiv() { Val = "0" };
            BottomMarginDiv bottomMarginDiv2 = new BottomMarginDiv() { Val = "0" };

            DivBorder divBorder2 = new DivBorder();
            TopBorder topBorder7 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder8 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder2.Append(topBorder7);
            divBorder2.Append(leftBorder3);
            divBorder2.Append(bottomBorder8);
            divBorder2.Append(rightBorder3);

            div2.Append(bodyDiv2);
            div2.Append(leftMarginDiv2);
            div2.Append(rightMarginDiv2);
            div2.Append(topMarginDiv2);
            div2.Append(bottomMarginDiv2);
            div2.Append(divBorder2);

            divs1.Append(div1);
            divs1.Append(div2);
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();
            RelyOnVML relyOnVML1 = new RelyOnVML();
            AllowPNG allowPNG1 = new AllowPNG();

            webSettings1.Append(divs1);
            webSettings1.Append(optimizeForBrowser1);
            webSettings1.Append(relyOnVML1);
            webSettings1.Append(allowPNG1);

            webSettingsPart1.WebSettings = webSettings1;
        }

        // Generates content of footerPart1.
        private void GenerateFooterPart1Content(FooterPart footerPart1)
        {
            Footer footer1 = new Footer();
            footer1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            Paragraph paragraph91 = new Paragraph() { RsidParagraphAddition = "00416BCF", RsidParagraphProperties = "00D30D81", RsidRunAdditionDefault = "00416BCF" };

            ParagraphProperties paragraphProperties78 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "a5" };
            FrameProperties frameProperties1 = new FrameProperties() { Wrap = TextWrappingValues.Around, HorizontalPosition = HorizontalAnchorValues.Margin, VerticalPosition = VerticalAnchorValues.Text, XAlign = HorizontalAlignmentValues.Center, Y = "1" };

            ParagraphMarkRunProperties paragraphMarkRunProperties78 = new ParagraphMarkRunProperties();
            RunStyle runStyle4 = new RunStyle() { Val = "a6" };

            paragraphMarkRunProperties78.Append(runStyle4);

            paragraphProperties78.Append(paragraphStyleId3);
            paragraphProperties78.Append(frameProperties1);
            paragraphProperties78.Append(paragraphMarkRunProperties78);

            paragraph91.Append(paragraphProperties78);

            Paragraph paragraph92 = new Paragraph() { RsidParagraphAddition = "00416BCF", RsidRunAdditionDefault = "00416BCF" };

            ParagraphProperties paragraphProperties79 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "a5" };

            paragraphProperties79.Append(paragraphStyleId4);

            paragraph92.Append(paragraphProperties79);

            footer1.Append(paragraph91);
            footer1.Append(paragraph92);

            footerPart1.Footer = footer1;
        }

        // Generates content of documentSettingsPart1.
        private void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings();
            settings1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            settings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            settings1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            settings1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            settings1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            settings1.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
            Zoom zoom1 = new Zoom() { Percent = "120" };
            ActiveWritingStyle activeWritingStyle1 = new ActiveWritingStyle() { Language = "ru-RU", VendorID = (UInt16Value)1U, DllVersion = 512, CheckStyle = true, ApplicationName = "MSWord" };
            ProofState proofState1 = new ProofState() { Spelling = ProofingStateValues.Clean, Grammar = ProofingStateValues.Clean };
            AttachedTemplate attachedTemplate1 = new AttachedTemplate() { Id = "rId1" };
            StylePaneFormatFilter stylePaneFormatFilter1 = new StylePaneFormatFilter() { Val = "3F01" };
            DocumentProtection documentProtection1 = new DocumentProtection() { Edit = DocumentProtectionValues.ReadOnly, Enforcement = false };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 708 };
            HyphenationZone hyphenationZone1 = new HyphenationZone() { Val = "357" };
            DoNotHyphenateCaps doNotHyphenateCaps1 = new DoNotHyphenateCaps();
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.DoNotCompress };

            FootnoteDocumentWideProperties footnoteDocumentWideProperties1 = new FootnoteDocumentWideProperties();
            FootnoteSpecialReference footnoteSpecialReference1 = new FootnoteSpecialReference() { Id = -1 };
            FootnoteSpecialReference footnoteSpecialReference2 = new FootnoteSpecialReference() { Id = 0 };

            footnoteDocumentWideProperties1.Append(footnoteSpecialReference1);
            footnoteDocumentWideProperties1.Append(footnoteSpecialReference2);

            EndnoteDocumentWideProperties endnoteDocumentWideProperties1 = new EndnoteDocumentWideProperties();
            EndnoteSpecialReference endnoteSpecialReference1 = new EndnoteSpecialReference() { Id = -1 };
            EndnoteSpecialReference endnoteSpecialReference2 = new EndnoteSpecialReference() { Id = 0 };

            endnoteDocumentWideProperties1.Append(endnoteSpecialReference1);
            endnoteDocumentWideProperties1.Append(endnoteSpecialReference2);
            Compatibility compatibility1 = new Compatibility();

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "00DF3879" };
            Rsid rsid11 = new Rsid() { Val = "00002973" };
            Rsid rsid12 = new Rsid() { Val = "00010A5E" };
            Rsid rsid13 = new Rsid() { Val = "000151D3" };
            Rsid rsid14 = new Rsid() { Val = "00020F5C" };
            Rsid rsid15 = new Rsid() { Val = "00034E0D" };
            Rsid rsid16 = new Rsid() { Val = "0003606F" };
            Rsid rsid17 = new Rsid() { Val = "0004432E" };
            Rsid rsid18 = new Rsid() { Val = "00054FAA" };
            Rsid rsid19 = new Rsid() { Val = "0006421F" };
            Rsid rsid20 = new Rsid() { Val = "0006505C" };
            Rsid rsid21 = new Rsid() { Val = "00072485" };
            Rsid rsid22 = new Rsid() { Val = "00076204" };
            Rsid rsid23 = new Rsid() { Val = "0009328A" };
            Rsid rsid24 = new Rsid() { Val = "000A4A48" };
            Rsid rsid25 = new Rsid() { Val = "000B0C3B" };
            Rsid rsid26 = new Rsid() { Val = "000B239B" };
            Rsid rsid27 = new Rsid() { Val = "000B3B96" };
            Rsid rsid28 = new Rsid() { Val = "000C00E1" };
            Rsid rsid29 = new Rsid() { Val = "000C1EF7" };
            Rsid rsid30 = new Rsid() { Val = "000C6B6F" };
            Rsid rsid31 = new Rsid() { Val = "000E4EB3" };
            Rsid rsid32 = new Rsid() { Val = "001124F4" };
            Rsid rsid33 = new Rsid() { Val = "001222E2" };
            Rsid rsid34 = new Rsid() { Val = "00140355" };
            Rsid rsid35 = new Rsid() { Val = "0014648A" };
            Rsid rsid36 = new Rsid() { Val = "00150214" };
            Rsid rsid37 = new Rsid() { Val = "00164FF0" };
            Rsid rsid38 = new Rsid() { Val = "001671F4" };
            Rsid rsid39 = new Rsid() { Val = "001A1288" };
            Rsid rsid40 = new Rsid() { Val = "001D1595" };
            Rsid rsid41 = new Rsid() { Val = "001E25EC" };
            Rsid rsid42 = new Rsid() { Val = "001E285D" };
            Rsid rsid43 = new Rsid() { Val = "001E2D95" };
            Rsid rsid44 = new Rsid() { Val = "001E2E89" };
            Rsid rsid45 = new Rsid() { Val = "001E4C83" };
            Rsid rsid46 = new Rsid() { Val = "001E5A99" };
            Rsid rsid47 = new Rsid() { Val = "001F1CBD" };
            Rsid rsid48 = new Rsid() { Val = "00214284" };
            Rsid rsid49 = new Rsid() { Val = "00222B39" };
            Rsid rsid50 = new Rsid() { Val = "00222D9A" };
            Rsid rsid51 = new Rsid() { Val = "002417C0" };
            Rsid rsid52 = new Rsid() { Val = "00250776" };
            Rsid rsid53 = new Rsid() { Val = "00256382" };
            Rsid rsid54 = new Rsid() { Val = "002566E7" };
            Rsid rsid55 = new Rsid() { Val = "002606A4" };
            Rsid rsid56 = new Rsid() { Val = "002649C4" };
            Rsid rsid57 = new Rsid() { Val = "00274353" };
            Rsid rsid58 = new Rsid() { Val = "002A5A9B" };
            Rsid rsid59 = new Rsid() { Val = "002B2A0B" };
            Rsid rsid60 = new Rsid() { Val = "002B7231" };
            Rsid rsid61 = new Rsid() { Val = "002D3437" };
            Rsid rsid62 = new Rsid() { Val = "002E648C" };
            Rsid rsid63 = new Rsid() { Val = "002F0859" };
            Rsid rsid64 = new Rsid() { Val = "002F114F" };
            Rsid rsid65 = new Rsid() { Val = "00311FAA" };
            Rsid rsid66 = new Rsid() { Val = "00312EF7" };
            Rsid rsid67 = new Rsid() { Val = "003179CD" };
            Rsid rsid68 = new Rsid() { Val = "00337A7B" };
            Rsid rsid69 = new Rsid() { Val = "00345646" };
            Rsid rsid70 = new Rsid() { Val = "003516C1" };
            Rsid rsid71 = new Rsid() { Val = "00353AF5" };
            Rsid rsid72 = new Rsid() { Val = "00396835" };
            Rsid rsid73 = new Rsid() { Val = "003B1AAE" };
            Rsid rsid74 = new Rsid() { Val = "003C42C4" };
            Rsid rsid75 = new Rsid() { Val = "003E6BA2" };
            Rsid rsid76 = new Rsid() { Val = "003F15A7" };
            Rsid rsid77 = new Rsid() { Val = "003F5C51" };
            Rsid rsid78 = new Rsid() { Val = "00416BCF" };
            Rsid rsid79 = new Rsid() { Val = "004178D7" };
            Rsid rsid80 = new Rsid() { Val = "00445C6E" };
            Rsid rsid81 = new Rsid() { Val = "00464C92" };
            Rsid rsid82 = new Rsid() { Val = "0047268D" };
            Rsid rsid83 = new Rsid() { Val = "004732E1" };
            Rsid rsid84 = new Rsid() { Val = "00486BDB" };
            Rsid rsid85 = new Rsid() { Val = "00494D22" };
            Rsid rsid86 = new Rsid() { Val = "004C67D2" };
            Rsid rsid87 = new Rsid() { Val = "004D5E1F" };
            Rsid rsid88 = new Rsid() { Val = "004D6D95" };
            Rsid rsid89 = new Rsid() { Val = "004E15AD" };
            Rsid rsid90 = new Rsid() { Val = "004E6F6C" };
            Rsid rsid91 = new Rsid() { Val = "004E7B44" };
            Rsid rsid92 = new Rsid() { Val = "0050283C" };
            Rsid rsid93 = new Rsid() { Val = "00516063" };
            Rsid rsid94 = new Rsid() { Val = "00520064" };
            Rsid rsid95 = new Rsid() { Val = "00527607" };
            Rsid rsid96 = new Rsid() { Val = "00527A5E" };
            Rsid rsid97 = new Rsid() { Val = "00534AE5" };
            Rsid rsid98 = new Rsid() { Val = "00535C3E" };
            Rsid rsid99 = new Rsid() { Val = "00545628" };
            Rsid rsid100 = new Rsid() { Val = "00551DC1" };
            Rsid rsid101 = new Rsid() { Val = "0056457F" };
            Rsid rsid102 = new Rsid() { Val = "0056665E" };
            Rsid rsid103 = new Rsid() { Val = "00572FE7" };
            Rsid rsid104 = new Rsid() { Val = "00574A0E" };
            Rsid rsid105 = new Rsid() { Val = "00576092" };
            Rsid rsid106 = new Rsid() { Val = "00581A80" };
            Rsid rsid107 = new Rsid() { Val = "005B16A0" };
            Rsid rsid108 = new Rsid() { Val = "005D56D5" };
            Rsid rsid109 = new Rsid() { Val = "005F1F99" };
            Rsid rsid110 = new Rsid() { Val = "005F6E47" };
            Rsid rsid111 = new Rsid() { Val = "00600802" };
            Rsid rsid112 = new Rsid() { Val = "006026C2" };
            Rsid rsid113 = new Rsid() { Val = "006029F9" };
            Rsid rsid114 = new Rsid() { Val = "006063B7" };
            Rsid rsid115 = new Rsid() { Val = "00611232" };
            Rsid rsid116 = new Rsid() { Val = "00615F2E" };
            Rsid rsid117 = new Rsid() { Val = "00624820" };
            Rsid rsid118 = new Rsid() { Val = "00624CA9" };
            Rsid rsid119 = new Rsid() { Val = "00633675" };
            Rsid rsid120 = new Rsid() { Val = "00650B2C" };
            Rsid rsid121 = new Rsid() { Val = "0065437A" };
            Rsid rsid122 = new Rsid() { Val = "00654D13" };
            Rsid rsid123 = new Rsid() { Val = "00656DEF" };
            Rsid rsid124 = new Rsid() { Val = "00670410" };
            Rsid rsid125 = new Rsid() { Val = "006752EE" };
            Rsid rsid126 = new Rsid() { Val = "00682122" };
            Rsid rsid127 = new Rsid() { Val = "00682F1A" };
            Rsid rsid128 = new Rsid() { Val = "00684C89" };
            Rsid rsid129 = new Rsid() { Val = "00691D5B" };
            Rsid rsid130 = new Rsid() { Val = "0069464F" };
            Rsid rsid131 = new Rsid() { Val = "00697C30" };
            Rsid rsid132 = new Rsid() { Val = "006A23C2" };
            Rsid rsid133 = new Rsid() { Val = "006A39C4" };
            Rsid rsid134 = new Rsid() { Val = "006A7A5E" };
            Rsid rsid135 = new Rsid() { Val = "006B2370" };
            Rsid rsid136 = new Rsid() { Val = "006B52CC" };
            Rsid rsid137 = new Rsid() { Val = "006F4ED3" };
            Rsid rsid138 = new Rsid() { Val = "00700208" };
            Rsid rsid139 = new Rsid() { Val = "007252A6" };
            Rsid rsid140 = new Rsid() { Val = "0077191F" };
            Rsid rsid141 = new Rsid() { Val = "00775C7F" };
            Rsid rsid142 = new Rsid() { Val = "00782905" };
            Rsid rsid143 = new Rsid() { Val = "00786295" };
            Rsid rsid144 = new Rsid() { Val = "00791026" };
            Rsid rsid145 = new Rsid() { Val = "00796E2B" };
            Rsid rsid146 = new Rsid() { Val = "007A37A8" };
            Rsid rsid147 = new Rsid() { Val = "007B17DA" };
            Rsid rsid148 = new Rsid() { Val = "007B54FC" };
            Rsid rsid149 = new Rsid() { Val = "007D2BB2" };
            Rsid rsid150 = new Rsid() { Val = "007F5B6A" };
            Rsid rsid151 = new Rsid() { Val = "0082407A" };
            Rsid rsid152 = new Rsid() { Val = "00854F10" };
            Rsid rsid153 = new Rsid() { Val = "00881B25" };
            Rsid rsid154 = new Rsid() { Val = "00885810" };
            Rsid rsid155 = new Rsid() { Val = "008952AA" };
            Rsid rsid156 = new Rsid() { Val = "00895753" };
            Rsid rsid157 = new Rsid() { Val = "008A1ECA" };
            Rsid rsid158 = new Rsid() { Val = "008A46A1" };
            Rsid rsid159 = new Rsid() { Val = "008C0E3E" };
            Rsid rsid160 = new Rsid() { Val = "008D0F16" };
            Rsid rsid161 = new Rsid() { Val = "008D1AD7" };
            Rsid rsid162 = new Rsid() { Val = "008D1B98" };
            Rsid rsid163 = new Rsid() { Val = "008E0776" };
            Rsid rsid164 = new Rsid() { Val = "008E46EB" };
            Rsid rsid165 = new Rsid() { Val = "008F0263" };
            Rsid rsid166 = new Rsid() { Val = "00913682" };
            Rsid rsid167 = new Rsid() { Val = "0092135F" };
            Rsid rsid168 = new Rsid() { Val = "009316BD" };
            Rsid rsid169 = new Rsid() { Val = "009345F5" };
            Rsid rsid170 = new Rsid() { Val = "00935E8D" };
            Rsid rsid171 = new Rsid() { Val = "009572B2" };
            Rsid rsid172 = new Rsid() { Val = "009809A9" };
            Rsid rsid173 = new Rsid() { Val = "00994E49" };
            Rsid rsid174 = new Rsid() { Val = "009952FB" };
            Rsid rsid175 = new Rsid() { Val = "009A487D" };
            Rsid rsid176 = new Rsid() { Val = "009A5348" };
            Rsid rsid177 = new Rsid() { Val = "009B1E4F" };
            Rsid rsid178 = new Rsid() { Val = "009C41A3" };
            Rsid rsid179 = new Rsid() { Val = "009D189B" };
            Rsid rsid180 = new Rsid() { Val = "009D724C" };
            Rsid rsid181 = new Rsid() { Val = "009E6299" };
            Rsid rsid182 = new Rsid() { Val = "009F2DB9" };
            Rsid rsid183 = new Rsid() { Val = "009F6F2A" };
            Rsid rsid184 = new Rsid() { Val = "009F780A" };
            Rsid rsid185 = new Rsid() { Val = "009F7FAE" };
            Rsid rsid186 = new Rsid() { Val = "00A01375" };
            Rsid rsid187 = new Rsid() { Val = "00A01B66" };
            Rsid rsid188 = new Rsid() { Val = "00A054B5" };
            Rsid rsid189 = new Rsid() { Val = "00A153C1" };
            Rsid rsid190 = new Rsid() { Val = "00A2593D" };
            Rsid rsid191 = new Rsid() { Val = "00A30E9C" };
            Rsid rsid192 = new Rsid() { Val = "00A33C10" };
            Rsid rsid193 = new Rsid() { Val = "00A56DA1" };
            Rsid rsid194 = new Rsid() { Val = "00A71FDE" };
            Rsid rsid195 = new Rsid() { Val = "00A723E4" };
            Rsid rsid196 = new Rsid() { Val = "00A82346" };
            Rsid rsid197 = new Rsid() { Val = "00A92423" };
            Rsid rsid198 = new Rsid() { Val = "00A95359" };
            Rsid rsid199 = new Rsid() { Val = "00A97AFC" };
            Rsid rsid200 = new Rsid() { Val = "00AA4282" };
            Rsid rsid201 = new Rsid() { Val = "00AA6CC3" };
            Rsid rsid202 = new Rsid() { Val = "00AD02C6" };
            Rsid rsid203 = new Rsid() { Val = "00AD2120" };
            Rsid rsid204 = new Rsid() { Val = "00AD2763" };
            Rsid rsid205 = new Rsid() { Val = "00AF22BD" };
            Rsid rsid206 = new Rsid() { Val = "00AF7427" };
            Rsid rsid207 = new Rsid() { Val = "00B00DA4" };
            Rsid rsid208 = new Rsid() { Val = "00B0580F" };
            Rsid rsid209 = new Rsid() { Val = "00B10ADB" };
            Rsid rsid210 = new Rsid() { Val = "00B33019" };
            Rsid rsid211 = new Rsid() { Val = "00B41D87" };
            Rsid rsid212 = new Rsid() { Val = "00B64EF4" };
            Rsid rsid213 = new Rsid() { Val = "00B70191" };
            Rsid rsid214 = new Rsid() { Val = "00B92CEE" };
            Rsid rsid215 = new Rsid() { Val = "00BB25E5" };
            Rsid rsid216 = new Rsid() { Val = "00BB3C66" };
            Rsid rsid217 = new Rsid() { Val = "00BB50A2" };
            Rsid rsid218 = new Rsid() { Val = "00BB58A0" };
            Rsid rsid219 = new Rsid() { Val = "00BC2F1A" };
            Rsid rsid220 = new Rsid() { Val = "00BC3916" };
            Rsid rsid221 = new Rsid() { Val = "00BC6DBF" };
            Rsid rsid222 = new Rsid() { Val = "00BF6BF5" };
            Rsid rsid223 = new Rsid() { Val = "00C204F3" };
            Rsid rsid224 = new Rsid() { Val = "00C31783" };
            Rsid rsid225 = new Rsid() { Val = "00C71E0B" };
            Rsid rsid226 = new Rsid() { Val = "00C80106" };
            Rsid rsid227 = new Rsid() { Val = "00C81D3A" };
            Rsid rsid228 = new Rsid() { Val = "00C957A3" };
            Rsid rsid229 = new Rsid() { Val = "00CA0175" };
            Rsid rsid230 = new Rsid() { Val = "00CA098C" };
            Rsid rsid231 = new Rsid() { Val = "00CA48EE" };
            Rsid rsid232 = new Rsid() { Val = "00CB0D6E" };
            Rsid rsid233 = new Rsid() { Val = "00CB3EE1" };
            Rsid rsid234 = new Rsid() { Val = "00CC3955" };
            Rsid rsid235 = new Rsid() { Val = "00CD0628" };
            Rsid rsid236 = new Rsid() { Val = "00CD3A75" };
            Rsid rsid237 = new Rsid() { Val = "00CE59D0" };
            Rsid rsid238 = new Rsid() { Val = "00CE7A3E" };
            Rsid rsid239 = new Rsid() { Val = "00CF3768" };
            Rsid rsid240 = new Rsid() { Val = "00D045BF" };
            Rsid rsid241 = new Rsid() { Val = "00D047AA" };
            Rsid rsid242 = new Rsid() { Val = "00D10361" };
            Rsid rsid243 = new Rsid() { Val = "00D11D1E" };
            Rsid rsid244 = new Rsid() { Val = "00D12358" };
            Rsid rsid245 = new Rsid() { Val = "00D20BF9" };
            Rsid rsid246 = new Rsid() { Val = "00D30977" };
            Rsid rsid247 = new Rsid() { Val = "00D30D81" };
            Rsid rsid248 = new Rsid() { Val = "00D34041" };
            Rsid rsid249 = new Rsid() { Val = "00D41B49" };
            Rsid rsid250 = new Rsid() { Val = "00D479F2" };
            Rsid rsid251 = new Rsid() { Val = "00D66F84" };
            Rsid rsid252 = new Rsid() { Val = "00D714F1" };
            Rsid rsid253 = new Rsid() { Val = "00D73787" };
            Rsid rsid254 = new Rsid() { Val = "00D82AD6" };
            Rsid rsid255 = new Rsid() { Val = "00D91BEA" };
            Rsid rsid256 = new Rsid() { Val = "00D9205A" };
            Rsid rsid257 = new Rsid() { Val = "00D924CD" };
            Rsid rsid258 = new Rsid() { Val = "00D9418E" };
            Rsid rsid259 = new Rsid() { Val = "00DA56AC" };
            Rsid rsid260 = new Rsid() { Val = "00DB7D17" };
            Rsid rsid261 = new Rsid() { Val = "00DC080F" };
            Rsid rsid262 = new Rsid() { Val = "00DC5151" };
            Rsid rsid263 = new Rsid() { Val = "00DC73BE" };
            Rsid rsid264 = new Rsid() { Val = "00DD344C" };
            Rsid rsid265 = new Rsid() { Val = "00DF3879" };
            Rsid rsid266 = new Rsid() { Val = "00DF61B7" };
            Rsid rsid267 = new Rsid() { Val = "00DF78B7" };
            Rsid rsid268 = new Rsid() { Val = "00E0522D" };
            Rsid rsid269 = new Rsid() { Val = "00E21F8B" };
            Rsid rsid270 = new Rsid() { Val = "00E447E6" };
            Rsid rsid271 = new Rsid() { Val = "00E52224" };
            Rsid rsid272 = new Rsid() { Val = "00E739AB" };
            Rsid rsid273 = new Rsid() { Val = "00E9125C" };
            Rsid rsid274 = new Rsid() { Val = "00E95B6A" };
            Rsid rsid275 = new Rsid() { Val = "00EA5149" };
            Rsid rsid276 = new Rsid() { Val = "00EA7A36" };
            Rsid rsid277 = new Rsid() { Val = "00EA7F81" };
            Rsid rsid278 = new Rsid() { Val = "00ED6913" };
            Rsid rsid279 = new Rsid() { Val = "00ED6DBD" };
            Rsid rsid280 = new Rsid() { Val = "00EE47EA" };
            Rsid rsid281 = new Rsid() { Val = "00EF5D56" };
            Rsid rsid282 = new Rsid() { Val = "00F011D0" };
            Rsid rsid283 = new Rsid() { Val = "00F0178B" };
            Rsid rsid284 = new Rsid() { Val = "00F118B5" };
            Rsid rsid285 = new Rsid() { Val = "00F3635D" };
            Rsid rsid286 = new Rsid() { Val = "00F45C0D" };
            Rsid rsid287 = new Rsid() { Val = "00F512CC" };
            Rsid rsid288 = new Rsid() { Val = "00F555BE" };
            Rsid rsid289 = new Rsid() { Val = "00F57005" };
            Rsid rsid290 = new Rsid() { Val = "00F57D30" };
            Rsid rsid291 = new Rsid() { Val = "00F651A7" };
            Rsid rsid292 = new Rsid() { Val = "00F65782" };
            Rsid rsid293 = new Rsid() { Val = "00F65818" };
            Rsid rsid294 = new Rsid() { Val = "00F738C5" };
            Rsid rsid295 = new Rsid() { Val = "00F76081" };
            Rsid rsid296 = new Rsid() { Val = "00FA132E" };
            Rsid rsid297 = new Rsid() { Val = "00FA4A4D" };
            Rsid rsid298 = new Rsid() { Val = "00FB65D1" };
            Rsid rsid299 = new Rsid() { Val = "00FC36E6" };
            Rsid rsid300 = new Rsid() { Val = "00FD0BFD" };
            Rsid rsid301 = new Rsid() { Val = "00FD73BD" };
            Rsid rsid302 = new Rsid() { Val = "00FE335F" };
            Rsid rsid303 = new Rsid() { Val = "00FE7382" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid11);
            rsids1.Append(rsid12);
            rsids1.Append(rsid13);
            rsids1.Append(rsid14);
            rsids1.Append(rsid15);
            rsids1.Append(rsid16);
            rsids1.Append(rsid17);
            rsids1.Append(rsid18);
            rsids1.Append(rsid19);
            rsids1.Append(rsid20);
            rsids1.Append(rsid21);
            rsids1.Append(rsid22);
            rsids1.Append(rsid23);
            rsids1.Append(rsid24);
            rsids1.Append(rsid25);
            rsids1.Append(rsid26);
            rsids1.Append(rsid27);
            rsids1.Append(rsid28);
            rsids1.Append(rsid29);
            rsids1.Append(rsid30);
            rsids1.Append(rsid31);
            rsids1.Append(rsid32);
            rsids1.Append(rsid33);
            rsids1.Append(rsid34);
            rsids1.Append(rsid35);
            rsids1.Append(rsid36);
            rsids1.Append(rsid37);
            rsids1.Append(rsid38);
            rsids1.Append(rsid39);
            rsids1.Append(rsid40);
            rsids1.Append(rsid41);
            rsids1.Append(rsid42);
            rsids1.Append(rsid43);
            rsids1.Append(rsid44);
            rsids1.Append(rsid45);
            rsids1.Append(rsid46);
            rsids1.Append(rsid47);
            rsids1.Append(rsid48);
            rsids1.Append(rsid49);
            rsids1.Append(rsid50);
            rsids1.Append(rsid51);
            rsids1.Append(rsid52);
            rsids1.Append(rsid53);
            rsids1.Append(rsid54);
            rsids1.Append(rsid55);
            rsids1.Append(rsid56);
            rsids1.Append(rsid57);
            rsids1.Append(rsid58);
            rsids1.Append(rsid59);
            rsids1.Append(rsid60);
            rsids1.Append(rsid61);
            rsids1.Append(rsid62);
            rsids1.Append(rsid63);
            rsids1.Append(rsid64);
            rsids1.Append(rsid65);
            rsids1.Append(rsid66);
            rsids1.Append(rsid67);
            rsids1.Append(rsid68);
            rsids1.Append(rsid69);
            rsids1.Append(rsid70);
            rsids1.Append(rsid71);
            rsids1.Append(rsid72);
            rsids1.Append(rsid73);
            rsids1.Append(rsid74);
            rsids1.Append(rsid75);
            rsids1.Append(rsid76);
            rsids1.Append(rsid77);
            rsids1.Append(rsid78);
            rsids1.Append(rsid79);
            rsids1.Append(rsid80);
            rsids1.Append(rsid81);
            rsids1.Append(rsid82);
            rsids1.Append(rsid83);
            rsids1.Append(rsid84);
            rsids1.Append(rsid85);
            rsids1.Append(rsid86);
            rsids1.Append(rsid87);
            rsids1.Append(rsid88);
            rsids1.Append(rsid89);
            rsids1.Append(rsid90);
            rsids1.Append(rsid91);
            rsids1.Append(rsid92);
            rsids1.Append(rsid93);
            rsids1.Append(rsid94);
            rsids1.Append(rsid95);
            rsids1.Append(rsid96);
            rsids1.Append(rsid97);
            rsids1.Append(rsid98);
            rsids1.Append(rsid99);
            rsids1.Append(rsid100);
            rsids1.Append(rsid101);
            rsids1.Append(rsid102);
            rsids1.Append(rsid103);
            rsids1.Append(rsid104);
            rsids1.Append(rsid105);
            rsids1.Append(rsid106);
            rsids1.Append(rsid107);
            rsids1.Append(rsid108);
            rsids1.Append(rsid109);
            rsids1.Append(rsid110);
            rsids1.Append(rsid111);
            rsids1.Append(rsid112);
            rsids1.Append(rsid113);
            rsids1.Append(rsid114);
            rsids1.Append(rsid115);
            rsids1.Append(rsid116);
            rsids1.Append(rsid117);
            rsids1.Append(rsid118);
            rsids1.Append(rsid119);
            rsids1.Append(rsid120);
            rsids1.Append(rsid121);
            rsids1.Append(rsid122);
            rsids1.Append(rsid123);
            rsids1.Append(rsid124);
            rsids1.Append(rsid125);
            rsids1.Append(rsid126);
            rsids1.Append(rsid127);
            rsids1.Append(rsid128);
            rsids1.Append(rsid129);
            rsids1.Append(rsid130);
            rsids1.Append(rsid131);
            rsids1.Append(rsid132);
            rsids1.Append(rsid133);
            rsids1.Append(rsid134);
            rsids1.Append(rsid135);
            rsids1.Append(rsid136);
            rsids1.Append(rsid137);
            rsids1.Append(rsid138);
            rsids1.Append(rsid139);
            rsids1.Append(rsid140);
            rsids1.Append(rsid141);
            rsids1.Append(rsid142);
            rsids1.Append(rsid143);
            rsids1.Append(rsid144);
            rsids1.Append(rsid145);
            rsids1.Append(rsid146);
            rsids1.Append(rsid147);
            rsids1.Append(rsid148);
            rsids1.Append(rsid149);
            rsids1.Append(rsid150);
            rsids1.Append(rsid151);
            rsids1.Append(rsid152);
            rsids1.Append(rsid153);
            rsids1.Append(rsid154);
            rsids1.Append(rsid155);
            rsids1.Append(rsid156);
            rsids1.Append(rsid157);
            rsids1.Append(rsid158);
            rsids1.Append(rsid159);
            rsids1.Append(rsid160);
            rsids1.Append(rsid161);
            rsids1.Append(rsid162);
            rsids1.Append(rsid163);
            rsids1.Append(rsid164);
            rsids1.Append(rsid165);
            rsids1.Append(rsid166);
            rsids1.Append(rsid167);
            rsids1.Append(rsid168);
            rsids1.Append(rsid169);
            rsids1.Append(rsid170);
            rsids1.Append(rsid171);
            rsids1.Append(rsid172);
            rsids1.Append(rsid173);
            rsids1.Append(rsid174);
            rsids1.Append(rsid175);
            rsids1.Append(rsid176);
            rsids1.Append(rsid177);
            rsids1.Append(rsid178);
            rsids1.Append(rsid179);
            rsids1.Append(rsid180);
            rsids1.Append(rsid181);
            rsids1.Append(rsid182);
            rsids1.Append(rsid183);
            rsids1.Append(rsid184);
            rsids1.Append(rsid185);
            rsids1.Append(rsid186);
            rsids1.Append(rsid187);
            rsids1.Append(rsid188);
            rsids1.Append(rsid189);
            rsids1.Append(rsid190);
            rsids1.Append(rsid191);
            rsids1.Append(rsid192);
            rsids1.Append(rsid193);
            rsids1.Append(rsid194);
            rsids1.Append(rsid195);
            rsids1.Append(rsid196);
            rsids1.Append(rsid197);
            rsids1.Append(rsid198);
            rsids1.Append(rsid199);
            rsids1.Append(rsid200);
            rsids1.Append(rsid201);
            rsids1.Append(rsid202);
            rsids1.Append(rsid203);
            rsids1.Append(rsid204);
            rsids1.Append(rsid205);
            rsids1.Append(rsid206);
            rsids1.Append(rsid207);
            rsids1.Append(rsid208);
            rsids1.Append(rsid209);
            rsids1.Append(rsid210);
            rsids1.Append(rsid211);
            rsids1.Append(rsid212);
            rsids1.Append(rsid213);
            rsids1.Append(rsid214);
            rsids1.Append(rsid215);
            rsids1.Append(rsid216);
            rsids1.Append(rsid217);
            rsids1.Append(rsid218);
            rsids1.Append(rsid219);
            rsids1.Append(rsid220);
            rsids1.Append(rsid221);
            rsids1.Append(rsid222);
            rsids1.Append(rsid223);
            rsids1.Append(rsid224);
            rsids1.Append(rsid225);
            rsids1.Append(rsid226);
            rsids1.Append(rsid227);
            rsids1.Append(rsid228);
            rsids1.Append(rsid229);
            rsids1.Append(rsid230);
            rsids1.Append(rsid231);
            rsids1.Append(rsid232);
            rsids1.Append(rsid233);
            rsids1.Append(rsid234);
            rsids1.Append(rsid235);
            rsids1.Append(rsid236);
            rsids1.Append(rsid237);
            rsids1.Append(rsid238);
            rsids1.Append(rsid239);
            rsids1.Append(rsid240);
            rsids1.Append(rsid241);
            rsids1.Append(rsid242);
            rsids1.Append(rsid243);
            rsids1.Append(rsid244);
            rsids1.Append(rsid245);
            rsids1.Append(rsid246);
            rsids1.Append(rsid247);
            rsids1.Append(rsid248);
            rsids1.Append(rsid249);
            rsids1.Append(rsid250);
            rsids1.Append(rsid251);
            rsids1.Append(rsid252);
            rsids1.Append(rsid253);
            rsids1.Append(rsid254);
            rsids1.Append(rsid255);
            rsids1.Append(rsid256);
            rsids1.Append(rsid257);
            rsids1.Append(rsid258);
            rsids1.Append(rsid259);
            rsids1.Append(rsid260);
            rsids1.Append(rsid261);
            rsids1.Append(rsid262);
            rsids1.Append(rsid263);
            rsids1.Append(rsid264);
            rsids1.Append(rsid265);
            rsids1.Append(rsid266);
            rsids1.Append(rsid267);
            rsids1.Append(rsid268);
            rsids1.Append(rsid269);
            rsids1.Append(rsid270);
            rsids1.Append(rsid271);
            rsids1.Append(rsid272);
            rsids1.Append(rsid273);
            rsids1.Append(rsid274);
            rsids1.Append(rsid275);
            rsids1.Append(rsid276);
            rsids1.Append(rsid277);
            rsids1.Append(rsid278);
            rsids1.Append(rsid279);
            rsids1.Append(rsid280);
            rsids1.Append(rsid281);
            rsids1.Append(rsid282);
            rsids1.Append(rsid283);
            rsids1.Append(rsid284);
            rsids1.Append(rsid285);
            rsids1.Append(rsid286);
            rsids1.Append(rsid287);
            rsids1.Append(rsid288);
            rsids1.Append(rsid289);
            rsids1.Append(rsid290);
            rsids1.Append(rsid291);
            rsids1.Append(rsid292);
            rsids1.Append(rsid293);
            rsids1.Append(rsid294);
            rsids1.Append(rsid295);
            rsids1.Append(rsid296);
            rsids1.Append(rsid297);
            rsids1.Append(rsid298);
            rsids1.Append(rsid299);
            rsids1.Append(rsid300);
            rsids1.Append(rsid301);
            rsids1.Append(rsid302);
            rsids1.Append(rsid303);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction();
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin8 = new M.LeftMargin() { Val = (UInt32Value)0U };
            M.RightMargin rightMargin8 = new M.RightMargin() { Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification() { Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent() { Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation() { Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin8);
            mathProperties1.Append(rightMargin8);
            mathProperties1.Append(defaultJustification1);
            mathProperties1.Append(wrapIndent1);
            mathProperties1.Append(integralLimitLocation1);
            mathProperties1.Append(naryLimitLocation1);
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages() { Val = "ru-RU" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping() { Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };
            DoNotIncludeSubdocsInStats doNotIncludeSubdocsInStats1 = new DoNotIncludeSubdocsInStats();

            ShapeDefaults shapeDefaults1 = new ShapeDefaults();

            Ovml.ShapeDefaults shapeDefaults2 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 11266 };
            Ovml.ColorMostRecentlyUsed colorMostRecentlyUsed1 = new Ovml.ColorMostRecentlyUsed() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Colors = "#5099dc,#498fe3" };

            shapeDefaults2.Append(colorMostRecentlyUsed1);

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "1" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults1.Append(shapeDefaults2);
            shapeDefaults1.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "," };
            ListSeparator listSeparator1 = new ListSeparator() { Val = ";" };

            settings1.Append(zoom1);
            settings1.Append(activeWritingStyle1);
            settings1.Append(proofState1);
            settings1.Append(attachedTemplate1);
            settings1.Append(stylePaneFormatFilter1);
            settings1.Append(documentProtection1);
            settings1.Append(defaultTabStop1);
            settings1.Append(hyphenationZone1);
            settings1.Append(doNotHyphenateCaps1);
            settings1.Append(characterSpacingControl1);
            settings1.Append(footnoteDocumentWideProperties1);
            settings1.Append(endnoteDocumentWideProperties1);
            settings1.Append(compatibility1);
            settings1.Append(rsids1);
            settings1.Append(mathProperties1);
            settings1.Append(themeFontLanguages1);
            settings1.Append(colorSchemeMapping1);
            settings1.Append(doNotIncludeSubdocsInStats1);
            settings1.Append(shapeDefaults1);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);

            documentSettingsPart1.Settings = settings1;
        }

        // Generates content of footerPart2.
        private void GenerateFooterPart2Content(FooterPart footerPart2)
        {
            Footer footer2 = new Footer();
            footer2.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer2.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer2.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer2.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer2.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer2.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer2.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            Paragraph paragraph93 = new Paragraph() { RsidParagraphAddition = "00416BCF", RsidParagraphProperties = "00D30D81", RsidRunAdditionDefault = "00BB25E5" };

            ParagraphProperties paragraphProperties80 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() { Val = "a5" };
            FrameProperties frameProperties2 = new FrameProperties() { Wrap = TextWrappingValues.Around, HorizontalPosition = HorizontalAnchorValues.Margin, VerticalPosition = VerticalAnchorValues.Text, XAlign = HorizontalAlignmentValues.Center, Y = "1" };

            ParagraphMarkRunProperties paragraphMarkRunProperties79 = new ParagraphMarkRunProperties();
            RunStyle runStyle5 = new RunStyle() { Val = "a6" };

            paragraphMarkRunProperties79.Append(runStyle5);

            paragraphProperties80.Append(paragraphStyleId5);
            paragraphProperties80.Append(frameProperties2);
            paragraphProperties80.Append(paragraphMarkRunProperties79);

            Run run72 = new Run();

            RunProperties runProperties68 = new RunProperties();
            RunStyle runStyle6 = new RunStyle() { Val = "a6" };

            runProperties68.Append(runStyle6);
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run72.Append(runProperties68);
            run72.Append(fieldChar1);

            Run run73 = new Run() { RsidRunAddition = "00416BCF" };

            RunProperties runProperties69 = new RunProperties();
            RunStyle runStyle7 = new RunStyle() { Val = "a6" };

            runProperties69.Append(runStyle7);
            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = "PAGE  ";

            run73.Append(runProperties69);
            run73.Append(fieldCode1);

            Run run74 = new Run();

            RunProperties runProperties70 = new RunProperties();
            RunStyle runStyle8 = new RunStyle() { Val = "a6" };

            runProperties70.Append(runStyle8);
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run74.Append(runProperties70);
            run74.Append(fieldChar2);

            paragraph93.Append(paragraphProperties80);
            paragraph93.Append(run72);
            paragraph93.Append(run73);
            paragraph93.Append(run74);

            Paragraph paragraph94 = new Paragraph() { RsidParagraphAddition = "00416BCF", RsidRunAdditionDefault = "00416BCF" };

            ParagraphProperties paragraphProperties81 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId6 = new ParagraphStyleId() { Val = "a5" };

            paragraphProperties81.Append(paragraphStyleId6);

            paragraph94.Append(paragraphProperties81);

            footer2.Append(paragraph93);
            footer2.Append(paragraph94);

            footerPart2.Footer = footer2;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "ken";
            document.PackageProperties.Title = "ОАО «ГАЗПРОМ»";
            document.PackageProperties.Revision = "4";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2014-04-13T13:36:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2014-04-17T13:40:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "ken";
            document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2011-12-29T05:02:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        #region Binary Data
        private string imagePart1Data = "iVBORw0KGgoAAAANSUhEUgAAAG8AAAC5BAMAAADUjHxhAAAAMFBMVEX//////wD/AP//AAAA//8A/wAAAP8AAAAAesP///8AAAAAAAAAAAAAAAAAAAAAAAAKCWfyAAAACnRSTlP///////////8AsswszwAAAAFiS0dEAIgFHUgAAAAMY21QUEpDbXAwNzEyAAAAB09tt6UAAAKfSURBVGje3drLeuMgDAXgPHqWR7tZHr3tpIkdZEBGwpfONyxq4vJbwiZtKvrQeKPp/3kkIKahzEECcxAHoMzAV6Ym5B3wx5Vcb4CvTMXkmoL2ScYhPlAmoCxfcpALxARUu1zDECuUNBRziEN+IdJwcwxDnAAlCWXbCUIaiCSsekGIU6CkoFTdFGQaLvnBZB2GbyOln4HMQ5h3cRKaWJ8XGYj1cD38Tm4elruTgDgE5S4oyxSTEAV+Qk9DyUAchJKDOg95FGJ5OhHI34Gre6H36/8Q8jhEBK63PwfNc8vAMi4F7bj3QMQgMAexC8WFmISYhJiEmIQMQWlg65Yf3SMIB3IAOQl7zkLxIAZQHcgh/BwaiCCsfwVwB67DyzkD4ULuQs5CDOGm94WchUhAicJSAmghZyFiUGvIFNQwpAsxCRmBOgF105ECcTdkHHLtBCGbK7whroNqoHwhL4bSQAQge2WZm6B8ISNQD0LaKlkMsi2v4VqopoYQgOuwFSIGDVuC2TIpQywBK/dGZYo+bNwH6gi27kdhDLXTXp9zxMCwq0v6wUQ3y8aF2m/b/Y5wwAB0XLWnEw+4aT0YcT0YCngqrG9lvz0fo4DXQ70byu1Qr4f8ZSj/MtQrYecKNdQOXC/GJCxJcAs5gP35jqGdNTPQucXP6iPZAHIW6jmQCUgPSgZqAuo5kHdAWrip890DeQKMPshn/cd1EsZzPQvGc0UFOQs1musO3A3JBgZzbWEwV7QwFJIdGApZxnRLiDKG0i9ajgM6ZVInpB3gFGaHAd1S8CigW3yWPdfZuNxZeFU6foG9ltV3dkr64px3/lkC9Qjn7GDbQtpT6sFBQUF8GKsojDeDenl6UMcusOHlPJ/xFpuzfMebel0W2UbsLtwX/Av0xBq+T5j9VwAAAABJRU5ErkJggg==";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion

    }
}
