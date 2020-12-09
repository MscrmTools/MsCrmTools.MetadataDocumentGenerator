using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using MsCrmTools.MetadataDocumentGenerator.Helper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Xml;

namespace MsCrmTools.MetadataDocumentGenerator.Generation
{
    public class WordDocumentOpenXml : IDocument
    {
        #region Variables

        private readonly List<EntityMetadata> emdCache;

        /// <summary>
        /// Word document
        /// </summary>
        private Document _innerDocument;

        /// <summary>
        /// Generation Settings
        /// </summary>
        private GenerationSettings _settings;

        private WordprocessingDocument _wordDocument = null;
        private IEnumerable<Entity> currentEntityForms;

        private BackgroundWorker worker;

        #endregion Variables

        #region Constructor

        /// <summary>
        /// Initializes a new instance of class WordDocument
        /// </summary>
        public WordDocumentOpenXml()
        {
            try
            {
                emdCache = new List<EntityMetadata>();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }

        #endregion Constructor

        public GenerationSettings Settings
        {
            get { return _settings; }
            set { _settings = value; }
        }

        public BackgroundWorker Worker
        {
            set { worker = value; }
        }

        #region Methods

        /// <summary>
        /// Add an attribute metadata
        /// </summary>
        /// <param name="attributeMetadataList">List of Attribute metadata</param>
        public void AddAttribute(IEnumerable<AttributeMetadata> attributeMetadataList)
        {
            var p = AddParagraph("Attributes", "Heading2", "Heading 2");

            var header = new List<string>
                             {
                                 "Logical Name",
                                 "Schema Name",
                                 "Display Name",
                                 "Attribute Type",
                                 "Description",
                                 "Is Custom",
                                 "Type"
                             };

            var amds = attributeMetadataList.OrderBy(attr => attr.SchemaName).Distinct(new AttributeMetadataComparer());

            var table = AddTable();

            int rowIndex = 0;

            foreach (var amd in amds)
            {
                var displayNameLabel = amd.DisplayName.LocalizedLabels.Count == 0
                                           ? null
                                           : amd.DisplayName.LocalizedLabels.FirstOrDefault(
                                               l => l.LanguageCode == _settings.DisplayNamesLangugageCode);
                var descriptionLabel = amd.Description.LocalizedLabels.Count == 0
                                           ? null
                                           : amd.Description.LocalizedLabels.FirstOrDefault(
                                               l => l.LanguageCode == _settings.DisplayNamesLangugageCode);

                var displayName = displayNameLabel != null ? displayNameLabel.Label : "Not Translated";
                var description = descriptionLabel != null ? descriptionLabel.Label : "Not Translated";

                var metadata = new List<string>
                {
                    amd.LogicalName,
                    amd.SchemaName,
                    displayName,
                    amd.AttributeType != null
                        ? amd.AttributeType.HasValue && amd.AttributeType == AttributeTypeCode.Virtual &&
                          amd is MultiSelectPicklistAttributeMetadata ? "MutliSelect OptionSet" :
                        amd.AttributeType.HasValue ? amd.AttributeType.Value.ToString() : string.Empty
                        : string.Empty,
                    description,
                    amd.IsCustomAttribute != null
                        ? amd.IsCustomAttribute.Value.ToString(CultureInfo.InvariantCulture)
                        : string.Empty,
                    (amd.SourceType ?? 0) == 0 ? "Simple" : (amd.SourceType ?? 0) == 1 ? "Calculated" : "Rollup"
                };

                if (_settings.AddRequiredLevelInformation)
                {
                    metadata.Add(amd.RequiredLevel.Value.ToString());
                    if (!header.Contains("Required Level")) header.Add("Required Level");
                }

                if (_settings.AddValidForAdvancedFind)
                {
                    metadata.Add(amd.IsValidForAdvancedFind.Value.ToString(CultureInfo.InvariantCulture));
                    if (!header.Contains("Valid for AF")) header.Add("Valid for AF");
                }

                if (_settings.AddAuditInformation)
                {
                    metadata.Add(amd.IsAuditEnabled.Value.ToString(CultureInfo.InvariantCulture));
                    if (!header.Contains("Audit Enabled")) header.Add("Audit Enabled");
                }

                if (_settings.AddFieldSecureInformation)
                {
                    metadata.Add(amd.IsSecured.Value.ToString(CultureInfo.InvariantCulture));
                    if (!header.Contains("Is Secured")) header.Add("Is Secured");
                }

                if (_settings.AddFormLocation)
                {
                    string data = string.Empty;

                    var entity = _settings.EntitiesToProceed.First(e => e.Name == amd.EntityLogicalName);

                    foreach (var form in entity.FormsDefinitions.Where(fd => entity.Forms.Contains(fd.Id) || entity.Forms.Count == 0))
                    {
                        var formName = form.GetAttributeValue<string>("name");
                        var xmlDocument = new XmlDocument();
                        xmlDocument.LoadXml(form["formxml"].ToString());

                        var controlNode = xmlDocument.SelectSingleNode("//control[@datafieldname='" + amd.LogicalName + "']");
                        if (controlNode != null)
                        {
                            XmlNodeList sectionNodes = controlNode.SelectNodes("ancestor::section");
                            XmlNodeList headerNodes = controlNode.SelectNodes("ancestor::header");
                            XmlNodeList footerNodes = controlNode.SelectNodes("ancestor::footer");

                            if (sectionNodes.Count > 0)
                            {
                                var sectionName = sectionNodes[0].SelectSingleNode("labels/label[@languagecode='" + _settings.DisplayNamesLangugageCode + "']").Attributes["description"].Value;

                                XmlNode tabNode = sectionNodes[0].SelectNodes("ancestor::tab")[0];
                                var tabName = tabNode.SelectSingleNode("labels/label[@languagecode='" + _settings.DisplayNamesLangugageCode + "']").Attributes["description"].Value;

                                if (data.Length > 0)
                                {
                                    data += "\n" + string.Format("{0}/{1}/{2}", formName, tabName, sectionName);
                                }
                                else
                                {
                                    data = string.Format("{0}/{1}/{2}", formName, tabName, sectionName);
                                }
                            }
                            else if (headerNodes.Count > 0)
                            {
                                if (data.Length > 0)
                                {
                                    data += "\n" + string.Format("{0}/Header", formName);
                                }
                                else
                                {
                                    data = string.Format("{0}/Header", formName);
                                }
                            }
                            else if (footerNodes.Count > 0)
                            {
                                if (data.Length > 0)
                                {
                                    data += "\n" + string.Format("{0}/Footer", formName);
                                }
                                else
                                {
                                    data = string.Format("{0}/Footer", formName);
                                }
                            }
                        }
                    }

                    metadata.Add(data);
                    if (!header.Contains("Form Location")) header.Add("Form Location");
                }

                metadata.Add(GetAddAdditionalData(amd));

                // add the header row now that they should have all been added
                if (rowIndex++ == 0)
                {
                    header.Add("Additional data");
                    AddHeaderRow(table, header);
                }
                // now add the new row with data
                AddTableRow(table, metadata);
            }
            _innerDocument.Body.InsertAfter(table, p);
        }

        /// <summary>
        /// Adds metadata of an entity
        /// </summary>
        /// <param name="emd">Entity metadata</param>
        public void AddEntityMetadata(EntityMetadata emd)
        {
            var displayNameLabel = emd.DisplayName.LocalizedLabels.Count == 0
                                       ? null
                                       : emd.DisplayName.LocalizedLabels.FirstOrDefault(
                                           l => l.LanguageCode == _settings.DisplayNamesLangugageCode);
            var pluralDisplayNameLabel = emd.DisplayCollectionName.LocalizedLabels.Count == 0
                                             ? null
                                             : emd.DisplayCollectionName.LocalizedLabels.FirstOrDefault(
                                                 l => l.LanguageCode == _settings.DisplayNamesLangugageCode);
            var descriptionLabel = emd.Description.LocalizedLabels.Count == 0
                                       ? null
                                       : emd.Description.LocalizedLabels.FirstOrDefault(
                                           l => l.LanguageCode == _settings.DisplayNamesLangugageCode);

            var displayName = displayNameLabel != null ? displayNameLabel.Label : "Not Translated";
            var pluralDisplayName = pluralDisplayNameLabel != null ? pluralDisplayNameLabel.Label : "Not Translated";
            var description = descriptionLabel != null ? descriptionLabel.Label : "Not Translated";

            AddParagraph("Entity: " + displayName, "Heading1", "Heading 1");

            AddParagraph("Metadata", "Heading2", "Heading 2");

            var table = AddTable();

            AddHeaderRow(table, new List<string> { "Property", "Value" });

            AddTableRow(table, new List<string> { "Display Name", displayName });
            AddTableRow(table, new List<string> { "Plural Display Name", pluralDisplayName });
            AddTableRow(table, new List<string> { "Description", description });
            AddTableRow(table, new List<string> { "Schema Name", emd.SchemaName });
            AddTableRow(table, new List<string> { "Logical Name", emd.LogicalName });
            AddTableRow(table, new List<string> {
                                         "Object Type Code",
                                         emd.ObjectTypeCode != null
                                             ? emd.ObjectTypeCode.Value.ToString(CultureInfo.InvariantCulture)
                                             : string.Empty
                                     });
            AddTableRow(table, new List<string> {
                                         "Is Custom Entity",
                                         (emd.IsCustomEntity != null && emd.IsCustomEntity.Value).ToString(
                                             CultureInfo.InvariantCulture)
                                     });
            AddTableRow(table, new List<string> {
                                         "Ownership Type",
                                         emd.OwnershipType != null ? emd.OwnershipType.Value.ToString() : string.Empty
                                     });

            var lastPara = _innerDocument.Body.ChildElements.Where(e => e is Paragraph).Last();
            _innerDocument.Body.InsertAfter(table, lastPara);
        }

        /// <summary>
        /// Generate the Word Document for the current selections
        /// </summary>
        /// <param name="service"></param>
        public void Generate(IOrganizationService service)
        {
            try
            {
                // create thew new document to which all content will be appended
                CreateWordprocessingDocument(Settings.FilePath);

                int totalEntities = _settings.EntitiesToProceed.Count;
                int processed = 0;

                foreach (var entity in _settings.EntitiesToProceed)
                {
                    ReportProgress(processed * 100 / totalEntities, string.Format("Processing entity '{0}'...", entity.Name));

                    var emd = emdCache.FirstOrDefault(x => x.LogicalName == entity.Name);
                    if (emd == null)
                    {
                        var reRequest = new RetrieveEntityRequest
                        {
                            LogicalName = entity.Name,
                            EntityFilters = EntityFilters.Entity | EntityFilters.Attributes
                        };
                        var reResponse = (RetrieveEntityResponse)service.Execute(reRequest);

                        emdCache.Add(reResponse.EntityMetadata);
                        emd = reResponse.EntityMetadata;
                    }

                    AddEntityMetadata(emd);

                    List<AttributeMetadata> amds = new List<AttributeMetadata>();

                    if (_settings.AddFormLocation)
                    {
                        currentEntityForms = MetadataHelper.RetrieveEntityFormList(emd.LogicalName, service);
                    }

                    switch (_settings.AttributesSelection)
                    {
                        case AttributeSelectionOption.AllAttributes:
                            amds = emd.Attributes.ToList();
                            break;

                        case AttributeSelectionOption.AttributesUnmanaged:
                            amds = emd.Attributes.Where(x => x.IsManaged.HasValue && x.IsManaged.Value == false).ToList();
                            break;

                        case AttributeSelectionOption.AttributesOptionSet:
                            amds =
                                emd.Attributes.Where(
                                    x => x.AttributeType != null && (x.AttributeType.Value == AttributeTypeCode.Boolean
                                                                     || x.AttributeType.Value == AttributeTypeCode.Picklist
                                                                     || x.AttributeType.Value == AttributeTypeCode.State
                                                                     || x.AttributeType.Value == AttributeTypeCode.Status
                                                                     || x.AttributeType == AttributeTypeCode.Virtual && x is MultiSelectPicklistAttributeMetadata)).ToList();

                            break;

                        case AttributeSelectionOption.AttributeManualySelected:
                            amds =
                                emd.Attributes.Where(
                                    x =>
                                    _settings.EntitiesToProceed.First(y => y.Name == emd.LogicalName).Attributes.Contains(
                                        x.LogicalName)).ToList();
                            break;

                        case AttributeSelectionOption.AttributesOnForm:

                            // If no forms selected, we search attributes in all forms
                            if (entity.Forms.Count == 0)
                            {
                                foreach (var form in entity.FormsDefinitions)
                                {
                                    var tempStringDoc = form.GetAttributeValue<string>("formxml");
                                    var tempDoc = new XmlDocument();
                                    tempDoc.LoadXml(tempStringDoc);

                                    amds.AddRange(emd.Attributes.Where(x =>
                                        tempDoc.SelectSingleNode("//control[@datafieldname='" + x.LogicalName + "']") !=
                                        null));
                                }
                            }
                            else
                            {
                                // else we parse selected forms
                                foreach (var formId in entity.Forms)
                                {
                                    var form = entity.FormsDefinitions.First(f => f.Id == formId);
                                    var tempStringDoc = form.GetAttributeValue<string>("formxml");
                                    var tempDoc = new XmlDocument();
                                    tempDoc.LoadXml(tempStringDoc);

                                    amds.AddRange(emd.Attributes.Where(x =>
                                        tempDoc.SelectSingleNode("//control[@datafieldname='" + x.LogicalName + "']") !=
                                        null));
                                }
                            }

                            break;

                        case AttributeSelectionOption.AttributesNotOnForm:
                            // If no forms selected, we search attributes in all forms
                            if (entity.Forms.Count == 0)
                            {
                                foreach (var form in entity.FormsDefinitions)
                                {
                                    var tempStringDoc = form.GetAttributeValue<string>("formxml");
                                    var tempDoc = new XmlDocument();
                                    tempDoc.LoadXml(tempStringDoc);

                                    amds.AddRange(emd.Attributes.Where(x =>
                                        tempDoc.SelectSingleNode("//control[@datafieldname='" + x.LogicalName + "']") ==
                                        null));
                                }
                            }
                            else
                            {
                                // else we parse selected forms
                                foreach (var formId in entity.Forms)
                                {
                                    var form = entity.FormsDefinitions.First(f => f.Id == formId);
                                    var tempStringDoc = form.GetAttributeValue<string>("formxml");
                                    var tempDoc = new XmlDocument();
                                    tempDoc.LoadXml(tempStringDoc);

                                    amds.AddRange(emd.Attributes.Where(x =>
                                        tempDoc.SelectSingleNode("//control[@datafieldname='" + x.LogicalName + "']") ==
                                        null));
                                }
                            }

                            break;
                    }

                    if (Settings.Prefixes != null && Settings.Prefixes.Count > 0)
                    {
                        var filteredAmds = new List<AttributeMetadata>();

                        foreach (var prefix in Settings.Prefixes)
                        {
                            filteredAmds.AddRange(amds.Where(a => a.LogicalName.StartsWith(prefix) /*|| a.IsCustomAttribute.Value == false*/));
                        }

                        amds = filteredAmds;
                    }

                    AddAttribute(amds);
                    processed++;
                }

                // before we save, set the orientation

                var sectionProps = new SectionProperties() { RsidR = "00F0206D", RsidSect = "0046795E" };
                var pageSize = new PageSize() { Width = (UInt32Value)24480U, Height = (UInt32Value)15840U, Orient = PageOrientationValues.Landscape, Code = (UInt16Value)17U };
                var pageMargin = new PageMargin() { Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
                var columns = new Columns() { Space = "720" };
                var docGrid = new DocGrid() { LinePitch = 360 };

                sectionProps.Append(pageSize);
                sectionProps.Append(pageMargin);
                sectionProps.Append(columns);
                sectionProps.Append(docGrid);

                _innerDocument.Body.Append(sectionProps);

                SaveDocument(_settings.FilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error has occurred attempting to generate your Word document:\n" + ex.Message);
            }
            finally
            {
                // ensure some cleanup!
                _wordDocument?.Close();
            }
        }

        /// <summary>
        /// Saves the current workbook
        /// </summary>
        /// <param name="path">Path where to save the document</param>
        public void SaveDocument(string path)
        {
            _innerDocument?.Save();
        }

        /// <summary>
        /// Add a header row to the Table
        /// </summary>
        /// <param name="table"></param>
        /// <param name="headers"></param>
        private void AddHeaderRow(Table table, List<string> headers, int rowIndex = 1)
        {
            var row = AddTableRow(table, headers, rowIndex);

            // set the current row properties as table header
            if (row.TableRowProperties == null)
            {
                row.TableRowProperties = new TableRowProperties();
            }
            // make this a header that breaks across pages
            row.TableRowProperties.AppendChild(new TableHeader());
        }

        /// <summary>
        /// Add a new paragraph object with a style ID
        /// </summary>
        /// <param name="content"></param>
        /// <param name="styleName"></param>
        /// <returns></returns>
        private Paragraph AddParagraph(string content = null, string styleId = null, string styleName = null)
        {
            var para = _innerDocument.Body.AppendChild(new Paragraph());
            var run = para.AppendChild(new Run());
            run.AppendChild(new Text(content));

            if (styleId != null)
            {
                OpenXmlHelper.ApplyStyleToParagraph(_wordDocument, styleId, styleName, para);
            }
            return para;
        }

        /// <summary>
        /// Add a new table to the document
        /// </summary>
        /// <returns></returns>
        private Table AddTable()
        {
            // Create an empty table.
            var table = new Table();
            OpenXmlHelper.AddTableStyle(_wordDocument, "GridTable4-Accent1");

            var tableProps = new TableProperties();
            var tableStyle = new TableStyle() { Val = "GridTable4-Accent1" };
            var tableWidth = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            var tableLook = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

            tableProps.Append(tableStyle);
            tableProps.Append(tableWidth);
            tableProps.Append(tableLook);

            // Append the TableProperties object to the empty table.
            table.Append(tableProps);

            return table;
        }

        /// <summary>
        /// Add the header to the current TableRow
        /// </summary>
        /// <param name="tr"></param>
        /// <param name="headerText"></param>
        private TableCell AddTableCell(TableRow tr, string headerText)
        {
            var tc = new TableCell(
                new Paragraph(
                    new Run(
                        new Text(headerText))
                    )
                );
            tr.Append(tc);

            return tc;
        }

        /// <summary>
        /// Add a new Table Row with a list of strings for cell content
        /// </summary>
        /// <param name="table"></param>
        /// <param name="cellContent"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        private TableRow AddTableRow(Table table, List<string> cellContent, int? rowIndex = null, string cellStyle = null)
        {
            var row = new TableRow();
            // add each cell to the new row
            foreach (var cell in cellContent)
            {
                AddTableCell(row, cell);
            }

            // insert the new TableRow at the correct location
            if (rowIndex != null)
            {
                table.InsertAt(row, rowIndex.Value);
            }
            else
            {
                table.Append(row);
            }

            return row;
        }

        /// <summary>
        /// Helper method to init the new document instance
        /// </summary>
        /// <param name="filepath"></param>
        private void CreateWordprocessingDocument(string filepath)
        {
            // Create a document by supplying the filepath.
            _wordDocument = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document);

            // Add a main document part.
            MainDocumentPart mainPart = _wordDocument.AddMainDocumentPart();

            // Create the document structure and add some text.
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            _innerDocument = mainPart.Document;
        }

        private string GetAddAdditionalData(AttributeMetadata amd)
        {
            if (amd.AttributeType != null)
                switch (amd.AttributeType.Value)
                {
                    case AttributeTypeCode.BigInt:
                        {
                            var bamd = (BigIntAttributeMetadata)amd;

                            return string.Format(
                                "Minimum value: {0}\nMaximum value: {1}",
                                bamd.MinValue.HasValue
                                    ? bamd.MinValue.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A",
                                bamd.MaxValue.HasValue
                                    ? bamd.MaxValue.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A");
                        }
                    case AttributeTypeCode.Boolean:
                        {
                            var bamd = (BooleanAttributeMetadata)amd;

                            var trueLabel = bamd.OptionSet.TrueOption.Label.LocalizedLabels.Count == 0
                                                ? null
                                                : bamd.OptionSet.TrueOption.Label.LocalizedLabels.FirstOrDefault(
                                                    l => l.LanguageCode == _settings.DisplayNamesLangugageCode);
                            var falseLabel = bamd.OptionSet.FalseOption.Label.LocalizedLabels.Count == 0
                                                 ? null
                                                 : bamd.OptionSet.FalseOption.Label.LocalizedLabels.
                                                        FirstOrDefault(
                                                            l =>
                                                            l.LanguageCode == _settings.DisplayNamesLangugageCode);

                            return string.Format(
                                "True: {0}\nFalse: {1}\nDefault Value: {2}",
                                bamd.OptionSet.TrueOption == null
                                    ? "N/A"
                                    : trueLabel != null ? trueLabel.Label : "Not Translated",
                                bamd.OptionSet.FalseOption == null
                                    ? "N/A"
                                    : falseLabel != null ? falseLabel.Label : "Not Translated",
                                (bamd.DefaultValue != null && bamd.DefaultValue.Value).ToString(
                                    CultureInfo.InvariantCulture));
                        }

                    case AttributeTypeCode.DateTime:
                        {
                            var damd = (DateTimeAttributeMetadata)amd;

                            return string.Format(
                                "Format: {0}",
                                damd.Format.HasValue ? damd.Format.Value.ToString() : "N/A");
                        }
                    case AttributeTypeCode.Decimal:
                        {
                            var damd = (DecimalAttributeMetadata)amd;

                            return string.Format(
                                "Minimum value: {0}\nMaximum value: {1}\nPrecision: {2}",
                                damd.MinValue.HasValue
                                    ? damd.MinValue.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A",
                                damd.MaxValue.HasValue
                                    ? damd.MaxValue.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A",
                                damd.Precision.HasValue
                                    ? damd.Precision.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A");
                        }
                    case AttributeTypeCode.Double:
                        {
                            var damd = (DoubleAttributeMetadata)amd;

                            return string.Format(
                                "Minimum value: {0}\nMaximum value: {1}\nPrecision: {2}",
                                damd.MinValue.HasValue
                                    ? damd.MinValue.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A",
                                damd.MaxValue.HasValue
                                    ? damd.MaxValue.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A",
                                damd.Precision.HasValue
                                    ? damd.Precision.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A");
                        }
                    case AttributeTypeCode.EntityName:
                        {
                            // Do nothing
                        }
                        break;

                    case AttributeTypeCode.Integer:
                        {
                            var iamd = (IntegerAttributeMetadata)amd;

                            return string.Format(
                                "Minimum value: {0}\nMaximum value: {1}",
                                iamd.MinValue.HasValue
                                    ? iamd.MinValue.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A",
                                iamd.MaxValue.HasValue
                                    ? iamd.MaxValue.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A");
                        }
                    case AttributeTypeCode.Customer:
                    case AttributeTypeCode.Owner:
                    case AttributeTypeCode.Lookup:
                        {
                            var lamd = (LookupAttributeMetadata)amd;

                            return lamd.Targets.Aggregate("Targets:", (current, entity) => current + ("\n" + entity));
                        }
                    case AttributeTypeCode.Memo:
                        {
                            var mamd = (MemoAttributeMetadata)amd;

                            return string.Format(
                                "Format: {0}\nMax length: {1}",
                                mamd.Format.HasValue ? mamd.Format.Value.ToString() : "N/A",
                                mamd.MaxLength.HasValue
                                    ? mamd.MaxLength.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A");
                        }
                    case AttributeTypeCode.Money:
                        {
                            var mamd = (MoneyAttributeMetadata)amd;

                            return string.Format(
                                "Minimum value: {0}\nMaximum value: {1}\nPrecision: {2}",
                                mamd.MinValue.HasValue
                                    ? mamd.MinValue.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A",
                                mamd.MaxValue.HasValue
                                    ? mamd.MaxValue.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A",
                                mamd.Precision.HasValue
                                    ? mamd.Precision.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A");
                        }

                    case AttributeTypeCode.PartyList:
                        {
                            // Do nothing
                        }
                        break;

                    case AttributeTypeCode.Virtual:
                        {
                            if (amd is MultiSelectPicklistAttributeMetadata mspamd)
                            {
                                var format = "Options:";

                                foreach (var omd in mspamd.OptionSet.Options)
                                {
                                    var optionLabel = omd.Label.LocalizedLabels.Count == 0
                                        ? null
                                        : omd.Label.LocalizedLabels.FirstOrDefault(
                                            l =>
                                                l.LanguageCode == _settings.DisplayNamesLangugageCode);

                                    format += string.Format("\n{0}: {1}",
                                        omd.Value,
                                        optionLabel != null ? optionLabel.Label : "Not Translated");
                                }

                                format += string.Format("\nDefault: {0}",
                                    mspamd.DefaultFormValue.HasValue
                                        ? mspamd.DefaultFormValue.Value.ToString(
                                            CultureInfo.InvariantCulture)
                                        : "N/A");

                                return format;
                            }

                            break;
                        }
                    case AttributeTypeCode.Picklist:
                        {
                            var pamd = (PicklistAttributeMetadata)amd;

                            var format = "Options:";

                            foreach (var omd in pamd.OptionSet.Options)
                            {
                                var optionLabel = omd.Label.LocalizedLabels.Count == 0
                                                      ? null
                                                      : omd.Label.LocalizedLabels.FirstOrDefault(
                                                          l =>
                                                          l.LanguageCode == _settings.DisplayNamesLangugageCode);

                                format += string.Format("\n{0}: {1}",
                                                        omd.Value,
                                                        optionLabel != null ? optionLabel.Label : "Not Translated");
                            }

                            format += string.Format("\nDefault: {0}",
                                                    pamd.DefaultFormValue.HasValue
                                                        ? pamd.DefaultFormValue.Value.ToString(
                                                            CultureInfo.InvariantCulture)
                                                        : "N/A");

                            return format;
                        }
                    case AttributeTypeCode.State:
                        {
                            var samd = (StateAttributeMetadata)amd;

                            var format = "States:";

                            foreach (var omd in samd.OptionSet.Options)
                            {
                                var optionLabel = omd.Label.LocalizedLabels.Count == 0
                                                      ? null
                                                      : omd.Label.LocalizedLabels.FirstOrDefault(
                                                          l =>
                                                          l.LanguageCode == _settings.DisplayNamesLangugageCode);

                                format += string.Format("\n{0}: {1}",
                                                        omd.Value,
                                                        optionLabel != null ? optionLabel.Label : "Not Translated");
                            }

                            return format;
                        }
                    case AttributeTypeCode.Status:
                        {
                            var samd = (StatusAttributeMetadata)amd;

                            string format = "States:";

                            foreach (var omd in samd.OptionSet.Options)
                            {
                                var optionLabel = omd.Label.LocalizedLabels.Count == 0
                                                      ? null
                                                      : omd.Label.LocalizedLabels.FirstOrDefault(
                                                          l =>
                                                          l.LanguageCode == _settings.DisplayNamesLangugageCode);

                                format += string.Format("\n{0}: {1}",
                                                        omd.Value,
                                                        optionLabel != null ? optionLabel.Label : "Not Translated");
                            }

                            return format;
                        }
                    case AttributeTypeCode.String:
                        {
                            var samd = (StringAttributeMetadata)amd;

                            return string.Format(
                                "Format: {0}\nMax length: {1}",
                                samd.Format.HasValue ? samd.Format.Value.ToString() : "N/A",
                                samd.MaxLength.HasValue
                                    ? samd.MaxLength.Value.ToString(CultureInfo.InvariantCulture)
                                    : "N/A");
                        }
                    case AttributeTypeCode.Uniqueidentifier:
                        {
                            // Do Nothing
                        }
                        break;
                }

            return string.Empty;
        }

        private void ReportProgress(int percentage, string message)
        {
            if (worker != null && worker.WorkerReportsProgress)
                worker.ReportProgress(percentage, message);
        }

        #endregion Methods
    }
}