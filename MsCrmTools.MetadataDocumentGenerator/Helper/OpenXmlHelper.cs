using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MsCrmTools.MetadataDocumentGenerator.Helper
{
    /// <summary>
    /// General helper class for OpenXml document generation
    /// TODO - move the hard wired styles to an Xml file or allow opening and copying from an existing Word Doc template 
    /// NOTE: The inline style code was done using the Document Reflector tool - better to ship a template, 
    /// open, copy styles from template.
    /// </summary>
    public class OpenXmlHelper
    {
        public static void AddTableStyle(WordprocessingDocument doc, string styleId) {
            // Get the Styles part for this document.
            var part = doc.MainDocumentPart.StyleDefinitionsPart;

            // If the Styles part does not exist, add it and then add the style.
            if (part == null)
            {
                part = AddStylesPartToPackage(doc);                
            }
            // If the style is not in the document, add it.
            if (!IsStyleIdInDocument(doc, styleId))
            {
                AddNewStyle(part, styleId);
            }
        }
        public static void ApplyStyleToParagraph(WordprocessingDocument doc, string styleid, string stylename, Paragraph p)
        {
            // If the paragraph has no ParagraphProperties object, create one.
            if (p.Elements<ParagraphProperties>().Count() == 0)
            {
                p.PrependChild<ParagraphProperties>(new ParagraphProperties());
            }

            // Get the paragraph properties element of the paragraph.
            var pPr = p.Elements<ParagraphProperties>().First();

            // Get the Styles part for this document.
            var part = doc.MainDocumentPart.StyleDefinitionsPart;

            // If the Styles part does not exist, add it and then add the style.
            if (part == null)
            {
                part = AddStylesPartToPackage(doc);
                AddNewStyle(part, styleid);
            }
            else
            {
                // If the style is not in the document, add it.
                if (IsStyleIdInDocument(doc, styleid) != true)
                {
                    // No match on styleid, so let's try style name.
                    string styleidFromName = GetStyleIdFromStyleName(doc, stylename);
                    if (styleidFromName == null)
                    {
                        AddNewStyle(part, styleid);
                    }
                    else
                        styleid = styleidFromName;
                }
            }

            // Set the style of the paragraph.
            pPr.ParagraphStyleId = new ParagraphStyleId() { Val = styleid };
        }

        // Return true if the style id is in the document, false otherwise.
        public static bool IsStyleIdInDocument(WordprocessingDocument doc,
            string styleid)
        {
            // Get access to the Styles element for this document.
            Styles s = doc.MainDocumentPart.StyleDefinitionsPart.Styles;

            // Check that there are styles and how many.
            int n = s.Elements<Style>().Count();
            if (n == 0)
                return false;

            // Look for a match on styleid.
            Style style = s.Elements<Style>()
                .Where(st => (st.StyleId == styleid) && 
                             ((st.Type == StyleValues.Paragraph) || (st.Type == StyleValues.Table))
                )
                .FirstOrDefault();
            if (style == null)
                return false;

            return true;
        }

        /// <summary>
        /// Return styleid that matches the styleName, or null when there's no match.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="styleName"></param>
        /// <returns></returns>
        public static string GetStyleIdFromStyleName(WordprocessingDocument doc, string styleName)
        {
            var stylePart = doc.MainDocumentPart.StyleDefinitionsPart;
            string styleId = stylePart.Styles.Descendants<StyleName>()
                .Where(s => s.Val.Value.Equals(styleName) &&
                    (((Style)s.Parent).Type == StyleValues.Paragraph))
                .Select(n => ((Style)n.Parent).StyleId).FirstOrDefault();
            return styleId;
        }

        /// <summary>
        /// Create a new style with the specified styleid and stylename and add it to the specified style definitions part.
        /// NOTE: this is only aware of Header 1 and Header 2 styles.  Need to be refactored to be more useful
        /// </summary>
        /// <param name="styleDefinitionsPart"></param>
        /// <param name="styleId"></param>
        /// <param name="stylename"></param>
        private static void AddNewStyle(StyleDefinitionsPart styleDefinitionsPart, string styleId)
        {
            // Get access to the root element of the styles part.
            Styles styles = styleDefinitionsPart.Styles;

            switch (styleId.ToLower())
            {
                case "heading1":
                    styles.Append(Generate_Heading1());
                    styles.Append(Generate_Heading1Char());
                    break;
                case "heading2":
                    styles.Append(Generate_Heading2());
                    styles.Append(Generate_Heading2Char());
                    break;

                case "gridtable4-accent1":
                    styles.Append(Generate_TableNormal());
                    styles.Append(Generate_GridTable4Accent1());
                    break;
            }
        }

        /// <summary>
        /// Add a StylesDefinitionsPart to the document.  Returns a reference to it.
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
        {
            StyleDefinitionsPart part;
            part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            Styles root = new Styles();
            root.Save(part);
            return part;
        }

        #region Generate specific styles by ID

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static Style Generate_Heading1()
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading1" };
            StyleName styleName1 = new StyleName() { Val = "heading 1" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading1Char" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid1 = new Rsid() { Val = "0052463A" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240", After = "0" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 0 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color1 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize1 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(primaryStyle1);
            style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }
            // Creates an Style instance and adds its children.
        public static Style Generate_Heading2()
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading2" };
            StyleName styleName1 = new StyleName() { Val = "heading 2" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading2Char" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid1 = new Rsid() { Val = "0052463A" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "40", After = "0" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 1 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color1 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize1 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(unhideWhenUsed1);
            style1.Append(primaryStyle1);
            style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }
        public static Style Generate_Heading2Char()
        {
            Style style1 = new Style() { Type = StyleValues.Character, StyleId = "Heading2Char", CustomStyle = true };
            StyleName styleName1 = new StyleName() { Val = "Heading 2 Char" };
            BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading2" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            Rsid rsid1 = new Rsid() { Val = "0052463A" };

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color1 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize1 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(rsid1);
            style1.Append(styleRunProperties1);
            return style1;
        }
        public static Style Generate_Heading1Char()
        {
            Style style1 = new Style() { Type = StyleValues.Character, StyleId = "Heading1Char", CustomStyle = true };
            StyleName styleName1 = new StyleName() { Val = "Heading 1 Char" };
            BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading1" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            Rsid rsid1 = new Rsid() { Val = "0052463A" };

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color1 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize1 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(rsid1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        // Creates an Style instance and adds its children.
        public static Style Generate_GridTable4Accent1()
        {
            Style style1 = new Style() { Type = StyleValues.Table, StyleId = "GridTable4-Accent1" };
            StyleName styleName1 = new StyleName() { Val = "Grid Table 4 Accent 1" };
            BasedOn basedOn1 = new BasedOn() { Val = "TableNormal" };
            UIPriority uIPriority1 = new UIPriority() { Val = 49 };
            Rsid rsid1 = new Rsid() { Val = "0052463A" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(spacingBetweenLines1);

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableStyleRowBandSize tableStyleRowBandSize1 = new TableStyleRowBandSize() { Val = 1 };
            TableStyleColumnBandSize tableStyleColumnBandSize1 = new TableStyleColumnBandSize() { Val = 1 };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "8EAADB", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "99", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);

            styleTableProperties1.Append(tableStyleRowBandSize1);
            styleTableProperties1.Append(tableStyleColumnBandSize1);
            styleTableProperties1.Append(tableBorders1);

            TableStyleProperties tableStyleProperties1 = new TableStyleProperties() { Type = TableStyleOverrideValues.FirstRow };

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            Color color1 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

            runPropertiesBaseStyle1.Append(bold1);
            runPropertiesBaseStyle1.Append(boldComplexScript1);
            runPropertiesBaseStyle1.Append(color1);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties1 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties1 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "4472C4", ThemeColor = ThemeColorValues.Accent1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "4472C4", ThemeColor = ThemeColorValues.Accent1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "4472C4", ThemeColor = ThemeColorValues.Accent1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "4472C4", ThemeColor = ThemeColorValues.Accent1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder2 = new InsideHorizontalBorder() { Val = BorderValues.Nil };
            InsideVerticalBorder insideVerticalBorder2 = new InsideVerticalBorder() { Val = BorderValues.Nil };

            tableCellBorders1.Append(topBorder2);
            tableCellBorders1.Append(leftBorder2);
            tableCellBorders1.Append(bottomBorder2);
            tableCellBorders1.Append(rightBorder2);
            tableCellBorders1.Append(insideHorizontalBorder2);
            tableCellBorders1.Append(insideVerticalBorder2);
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "4472C4", ThemeFill = ThemeColorValues.Accent1 };

            tableStyleConditionalFormattingTableCellProperties1.Append(tableCellBorders1);
            tableStyleConditionalFormattingTableCellProperties1.Append(shading1);

            tableStyleProperties1.Append(runPropertiesBaseStyle1);
            tableStyleProperties1.Append(tableStyleConditionalFormattingTableProperties1);
            tableStyleProperties1.Append(tableStyleConditionalFormattingTableCellProperties1);

            TableStyleProperties tableStyleProperties2 = new TableStyleProperties() { Type = TableStyleOverrideValues.LastRow };

            RunPropertiesBaseStyle runPropertiesBaseStyle2 = new RunPropertiesBaseStyle();
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();

            runPropertiesBaseStyle2.Append(bold2);
            runPropertiesBaseStyle2.Append(boldComplexScript2);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties2 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties2 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders2 = new TableCellBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Double, Color = "4472C4", ThemeColor = ThemeColorValues.Accent1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(topBorder3);

            tableStyleConditionalFormattingTableCellProperties2.Append(tableCellBorders2);

            tableStyleProperties2.Append(runPropertiesBaseStyle2);
            tableStyleProperties2.Append(tableStyleConditionalFormattingTableProperties2);
            tableStyleProperties2.Append(tableStyleConditionalFormattingTableCellProperties2);

            TableStyleProperties tableStyleProperties3 = new TableStyleProperties() { Type = TableStyleOverrideValues.FirstColumn };

            RunPropertiesBaseStyle runPropertiesBaseStyle3 = new RunPropertiesBaseStyle();
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();

            runPropertiesBaseStyle3.Append(bold3);
            runPropertiesBaseStyle3.Append(boldComplexScript3);

            tableStyleProperties3.Append(runPropertiesBaseStyle3);

            TableStyleProperties tableStyleProperties4 = new TableStyleProperties() { Type = TableStyleOverrideValues.LastColumn };

            RunPropertiesBaseStyle runPropertiesBaseStyle4 = new RunPropertiesBaseStyle();
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();

            runPropertiesBaseStyle4.Append(bold4);
            runPropertiesBaseStyle4.Append(boldComplexScript4);

            tableStyleProperties4.Append(runPropertiesBaseStyle4);

            TableStyleProperties tableStyleProperties5 = new TableStyleProperties() { Type = TableStyleOverrideValues.Band1Vertical };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties3 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties3 = new TableStyleConditionalFormattingTableCellProperties();
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D9E2F3", ThemeFill = ThemeColorValues.Accent1, ThemeFillTint = "33" };

            tableStyleConditionalFormattingTableCellProperties3.Append(shading2);

            tableStyleProperties5.Append(tableStyleConditionalFormattingTableProperties3);
            tableStyleProperties5.Append(tableStyleConditionalFormattingTableCellProperties3);

            TableStyleProperties tableStyleProperties6 = new TableStyleProperties() { Type = TableStyleOverrideValues.Band1Horizontal };
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties4 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties4 = new TableStyleConditionalFormattingTableCellProperties();
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D9E2F3", ThemeFill = ThemeColorValues.Accent1, ThemeFillTint = "33" };

            tableStyleConditionalFormattingTableCellProperties4.Append(shading3);

            tableStyleProperties6.Append(tableStyleConditionalFormattingTableProperties4);
            tableStyleProperties6.Append(tableStyleConditionalFormattingTableCellProperties4);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(uIPriority1);
            style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleTableProperties1);
            style1.Append(tableStyleProperties1);
            style1.Append(tableStyleProperties2);
            style1.Append(tableStyleProperties3);
            style1.Append(tableStyleProperties4);
            style1.Append(tableStyleProperties5);
            style1.Append(tableStyleProperties6);
            return style1;
        }
        public static Style Generate_TableNormal()
        {
            Style style1 = new Style() { Type = StyleValues.Table, StyleId = "TableNormal", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault1);

            style1.Append(styleName1);
            style1.Append(uIPriority1);
            style1.Append(semiHidden1);
            style1.Append(unhideWhenUsed1);
            style1.Append(styleTableProperties1);
            return style1;
        }
        #endregion
    }
    }
