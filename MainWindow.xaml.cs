// ===============================
// UPDATED BY: Liam Wright
// UPDATE DATE: 30/12/2015
// PURPOSE: Accept additional XML's
// ===============================
// Change History:
/*
 * Function CreateRATFile edited to read input from all XMLS
 * in the draft folder. Loop implemented to read multiple XML's.
 * 
 * Logic added to add all jobs from new XML's
 * 
 */
//==================================

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Markup;
using System.Xml.Linq;
using System.Xml.Schema;

namespace RATFileGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnPrevBrowse_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                folderBrowserDialog.Description = "Select Previous Version Location";
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    this.txtPrevVer.Text = folderBrowserDialog.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "RAT Generator", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
            finally
            {
            }
        }

        private void btnCurrBrowse_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                folderBrowserDialog.Description = "Select Previous Version Location";
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    this.txtCurrVer.Text = folderBrowserDialog.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "RAT Generator", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
            finally
            {
            }
        }

        private bool CreateRATFile(string RATFileName, DirectoryInfo currentLocalDirectory, DirectoryInfo previousLocalDirectory, string component)
        {
            /* ------------------------------------
             * Note: This section is to compare the folder contents
             *-------------------------------------
             */
            //Set the draft folder to Current Directory/DRAFTS
            DirectoryInfo draftFolder = new DirectoryInfo(currentLocalDirectory.FullName + "\\DRAFTS\\");
            DirectoryInfo prevDraftFolder = new DirectoryInfo(previousLocalDirectory.FullName + "\\DRAFTS\\");

            //Add all files into 'files' array.
            FileInfo[] files = currentLocalDirectory.GetFiles("*.*", SearchOption.AllDirectories);

            //Initialise List 'List'
            List<FileSystemInfo> list = new List<FileSystemInfo>();

            //Add full list of files in previous directory to list.
            list.AddRange(previousLocalDirectory.GetFiles("*.*", SearchOption.AllDirectories));

            //If a RAT File exists, then delete it.
            if (File.Exists(RATFileName))
            {
                File.Delete(RATFileName);
            }


            //Initialise Streamwriter to write to RATFileName file.
            StreamWriter RATFileWriter = new StreamWriter(RATFileName,false);

            //Add in Header
            RATFileWriter.WriteLine("Name, Type, Status, Current Modified, Previous Modified");

            //Set 'files' array to 'array' Array.
            FileInfo[] array = files;

            // Create null FileInfo variable: releaseFile
            FileInfo releaseFile;

            //As long as int i is less than the length of the array, loop.
            for (int i = 0; i < array.Length; i++)
            {
                //Set releaseFile to current entry in 'array' Array.
                releaseFile = array[i];

                //If CompareFile function returns a true, continue...
                if (this.CompareFile(releaseFile))
                {
                    FileSystemInfo fileSystemInfo = list.Find((FileSystemInfo prf) => prf.Name.Equals(releaseFile.Name));
                    if (fileSystemInfo != null && fileSystemInfo.Exists)
                    {
                        if (this.CalculateHash(releaseFile) != this.CalculateHash(fileSystemInfo))
                        {
                            RATFileWriter.WriteLine("{0}, File, Changed, {1}, {2}", releaseFile.Name, releaseFile.LastWriteTime, fileSystemInfo.LastWriteTime);
                        }
                    }
                    else
                    {
                        RATFileWriter.WriteLine("{0}, File, Added, {1}", releaseFile.Name, releaseFile.LastWriteTime);
                    }
                }
            }
            IEnumerable<FileSystemInfo> enumerable = list.Except(
                from previousVersionFile in list
                join currentFile in files on previousVersionFile.Name equals currentFile.Name
                select previousVersionFile);
            enumerable =
                from removedFile in enumerable
                where removedFile is FileInfo
                select removedFile;
            if (enumerable != null)
            {
                using (List<FileSystemInfo>.Enumerator enumerator = enumerable.ToList<FileSystemInfo>().GetEnumerator())
                {
                    while (enumerator.MoveNext())
                    {
                        FileInfo fileInfo = (FileInfo)enumerator.Current;
                        if (this.CompareFile(fileInfo))
                        {
                         removeJobs(fileInfo, prevDraftFolder, RATFileWriter);
                         RATFileWriter.WriteLine("{0}, File, Removed, , {1}", fileInfo.Name, fileInfo.LastWriteTime);
                        }
                    }
                }
            }

            /*---------------------------------------------------------------------------
             * Loop to accept all input from Draft folder XML's
             * --------------------------------------------------------------------------
             */
            


            //Add all xml files in DRAFT folder into FileInfo array
            FileInfo[] xmlFiles = draftFolder.GetFiles("*.xml");

            //For each file in xmlFile array, do the following...
            foreach (FileInfo xmlFile in xmlFiles)
            {

                string prevXML = xmlFile.Name;
                string prevXMLFullPath = prevDraftFolder + prevXML;
                //Add all prevXml files in DRAFT folder into FileInfo array
                FileInfo[] prevXmlFiles = prevDraftFolder.GetFiles("*.xml");
                         
                //set URI to xmlFile full name (Eg: CurrentDirectory\Draft\Something.xml)
                string uri = xmlFile.FullName;
                //---------------------------------------------------------------------------------------------------------------------------------------
                // Search through list and find fileName.xml (This is the previous version)
                FileSystemInfo fileSystemInfo2 = list.Find((FileSystemInfo file) => file.Name.ToUpper().Equals(prevXML.ToUpper()));

                
                if (fileSystemInfo2 != null && XDocument.Load(uri) != null)
                {
                    XDocument xDocument = XDocument.Load(prevXMLFullPath);
                    XDocument xDocument2 = XDocument.Load(uri);
                    IEnumerable<XElement> source =
                        from job in xDocument.Descendants("JOB")
                        select job;
                    List<XElement> nodeOldJobs = source.ToList<XElement>();
                    IEnumerable<XElement> source2 =
                        from job in xDocument2.Descendants("JOB")
                        select job;
                    List<XElement> list2 = source2.ToList<XElement>();
                    list2.ForEach(delegate(XElement newJob)
                    {
                        XElement xElement = nodeOldJobs.Find((XElement job) => job.Attribute("JOBNAME").Value.Equals(newJob.Attribute("JOBNAME").Value));
                        if (xElement != null)
                        {
                            XDocument doc = new XDocument(new object[]
						    {
							    new XElement(newJob)
						    });
                            XDocument doc2 = new XDocument(new object[]
						    {
							    new XElement(xElement)
						    });
                            if (!DeepEqualsWithNormalization(doc, doc2, null))
                            {
                                RATFileWriter.WriteLine("{0}, Job, Changed, ,", newJob.Attribute("JOBNAME").Value);
                            }
                            nodeOldJobs.Remove(xElement);
                        }
                        else
                        {
                            RATFileWriter.WriteLine("{0}, Job, Added, ,", newJob.Attribute("JOBNAME").Value);
                        }
                    });
                    nodeOldJobs.ForEach(delegate(XElement oldJob)
                    {
                        RATFileWriter.WriteLine("{0}, Job, Removed, ,", oldJob.Attribute("JOBNAME").Value);
                    });
              }
                //If there is a NEW XML, add all jobs to RAT File.
                if (fileSystemInfo2 == null && XDocument.Load(uri) != null)
                {
                    XDocument xDocument2 = XDocument.Load(uri);
                    IEnumerable<XElement> source2 =
                          from job in xDocument2.Descendants("JOB")
                          select job;
                    List<XElement> list2 = source2.ToList<XElement>();

                    list2.ForEach(delegate(XElement newJob)
                    {                 
                       RATFileWriter.WriteLine("{0}, Job, Added, ,", newJob.Attribute("JOBNAME").Value);
                    });
                }

            }
            RATFileWriter.Close();
            currentLocalDirectory = null;
            return true;
        }
        private XDocument Normalize(XDocument source, XmlSchemaSet schema)
        {
            bool havePSVI = false;
            if (schema != null)
            {
                source.Validate(schema, null, true);
                havePSVI = true;
            }
            return new XDocument(source.Declaration, new object[]
			{
				source.Nodes().Select(delegate(XNode n)
				{
					XNode result;
					if (n is XComment || n is XProcessingInstruction || n is XText)
					{
						result = null;
					}
					else
					{
						XElement xElement = n as XElement;
						if (xElement != null)
						{
							result = MainWindow.NormalizeElement(xElement, havePSVI);
						}
						else
						{
							result = n;
						}
					}
					return result;
				})
			});
        }
        private static IEnumerable<XAttribute> NormalizeAttributes(XElement element, bool havePSVI)
        {
            return (
                from a in element.Attributes()
                where !a.IsNamespaceDeclaration && a.Name != Xsi.schemaLocation && a.Name != Xsi.noNamespaceSchemaLocation
                orderby a.Name.NamespaceName, a.Name.LocalName
                select a).Select(delegate(XAttribute a)
                {
                    XAttribute result;
                    if (havePSVI)
                    {
                        XmlTypeCode typeCode = a.GetSchemaInfo().SchemaType.TypeCode;
                        XmlTypeCode xmlTypeCode = typeCode;
                        switch (xmlTypeCode)
                        {
                            case XmlTypeCode.Boolean:
                                result = new XAttribute(a.Name, (bool)a);
                                return result;
                            case XmlTypeCode.Decimal:
                                result = new XAttribute(a.Name, (decimal)a);
                                return result;
                            case XmlTypeCode.Float:
                                result = new XAttribute(a.Name, (float)a);
                                return result;
                            case XmlTypeCode.Double:
                                result = new XAttribute(a.Name, (double)a);
                                return result;
                            case XmlTypeCode.Duration:
                                break;
                            case XmlTypeCode.DateTime:
                                result = new XAttribute(a.Name, (DateTime)a);
                                return result;
                            default:
                                if (xmlTypeCode == XmlTypeCode.HexBinary || xmlTypeCode == XmlTypeCode.Language)
                                {
                                    result = new XAttribute(a.Name, ((string)a).ToLower());
                                    return result;
                                }
                                break;
                        }
                    }
                    if (a.Name.LocalName.Equals("CREATION_DATE") | a.Name.LocalName.Equals("CREATION_TIME") | a.Name.LocalName.Equals("CREATION_USER") | a.Name.LocalName.Equals("CHANGE_DATE") | a.Name.LocalName.Equals("CHANGE_TIME") | a.Name.LocalName.Equals("CHANGE_USERID") | a.Name.LocalName.Equals("AUTHOR") | a.Name.LocalName.Equals("PREVENTNCT2"))
                    {
                        result = null;
                    }
                    else
                    {
                        result = a;
                    }
                    return result;
                });
        }
        private static XNode NormalizeNode(XNode node, bool havePSVI)
        {
            XNode result;
            if (node is XComment || node is XProcessingInstruction)
            {
                result = null;
            }
            else
            {
                XElement xElement = node as XElement;
                if (xElement != null)
                {
                    result = MainWindow.NormalizeElement(xElement, havePSVI);
                }
                else
                {
                    result = node;
                }
            }
            return result;
        }
        private static XElement NormalizeElement(XElement element, bool havePSVI)
        {
            XElement result;
            if (havePSVI)
            {
                IXmlSchemaInfo schemaInfo = element.GetSchemaInfo();
                XmlTypeCode typeCode = schemaInfo.SchemaType.TypeCode;
                switch (typeCode)
                {
                    case XmlTypeCode.Boolean:
                        result = new XElement(element.Name, new object[]
					{
						MainWindow.NormalizeAttributes(element, havePSVI),
						(bool)element
					});
                        return result;
                    case XmlTypeCode.Decimal:
                        result = new XElement(element.Name, new object[]
					{
						MainWindow.NormalizeAttributes(element, havePSVI),
						(decimal)element
					});
                        return result;
                    case XmlTypeCode.Float:
                        result = new XElement(element.Name, new object[]
					{
						MainWindow.NormalizeAttributes(element, havePSVI),
						(float)element
					});
                        return result;
                    case XmlTypeCode.Double:
                        result = new XElement(element.Name, new object[]
					{
						MainWindow.NormalizeAttributes(element, havePSVI),
						(double)element
					});
                        return result;
                    case XmlTypeCode.Duration:
                        break;
                    case XmlTypeCode.DateTime:
                        result = new XElement(element.Name, new object[]
					{
						MainWindow.NormalizeAttributes(element, havePSVI),
						(DateTime)element
					});
                        return result;
                    default:
                        if (typeCode == XmlTypeCode.HexBinary || typeCode == XmlTypeCode.Language)
                        {
                            result = new XElement(element.Name, new object[]
						{
							MainWindow.NormalizeAttributes(element, havePSVI),
							((string)element).ToLower()
						});
                            return result;
                        }
                        break;
                }
                XName arg_218_0 = element.Name;
                object[] array = new object[2];
                array[0] = MainWindow.NormalizeAttributes(element, havePSVI);
                array[1] =
                    from n in element.Nodes()
                    select MainWindow.NormalizeNode(n, havePSVI);
                result = new XElement(arg_218_0, array);
            }
            else
            {
                XName arg_264_0 = element.Name;
                object[] array = new object[2];
                array[0] = MainWindow.NormalizeAttributes(element, havePSVI);
                array[1] =
                    from n in element.Nodes()
                    select MainWindow.NormalizeNode(n, havePSVI);
                result = new XElement(arg_264_0, array);
            }
            return result;
        }

        private bool DeepEqualsWithNormalization(XDocument doc1, XDocument doc2, XmlSchemaSet schemaSet)
        {
            XDocument n = this.Normalize(doc1, schemaSet);
            XDocument n2 = this.Normalize(doc2, schemaSet);
            return XNode.DeepEquals(n, n2);
        }

        //Function to ensure that we want to compare the file
        private bool CompareFile(FileInfo fileToCheck)
        {
            return !(fileToCheck.Name.ToUpper().Contains(".HTM") | fileToCheck.Name.ToUpper().Contains(".HTML") | fileToCheck.Name.ToUpper().Contains("AUTOEDIT") | fileToCheck.Name.ToUpper().Contains("CREATEOPINS") | fileToCheck.Name.ToUpper().Contains(".CSV") | fileToCheck.Name.ToUpper().Contains("CONTROLM.DLL") | fileToCheck.Name.ToUpper().Contains("CREATE_OPINS"));
        }

        private string CalculateHash(FileSystemInfo fileToHash)
        {
            StreamReader streamReader = new StreamReader(fileToHash.FullName);
            UnicodeEncoding unicodeEncoding = new UnicodeEncoding();
            byte[] bytes = unicodeEncoding.GetBytes(streamReader.ReadToEnd());
            MD5CryptoServiceProvider mD5CryptoServiceProvider = new MD5CryptoServiceProvider();
            byte[] inArray = mD5CryptoServiceProvider.ComputeHash(bytes);
            streamReader.Close();
            return Convert.ToBase64String(inArray);
        }
        
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }

        private void btnGenerate_Click(object sender, RoutedEventArgs e)
        {
            string component = "";
            try
            {
                if (!string.IsNullOrEmpty(this.txtPrevVer.Text) && !string.IsNullOrEmpty(this.txtCurrVer.Text))
                {
                    string[] array = Path.GetFileName(this.txtCurrVer.Text).Split(new char[]
					{
						'_'
					});
                    if (array.Length > 0)
                    {
                        component = array.GetValue(0).ToString();
                    }
                    if (this.CreateRATFile(this.txtRATFile.Text, new DirectoryInfo(this.txtCurrVer.Text), new DirectoryInfo(this.txtPrevVer.Text), component))
                    {
                        System.Windows.MessageBox.Show("RAT File created successfully", "Success", MessageBoxButton.OK, MessageBoxImage.Information, MessageBoxResult.OK);
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("Please Select the Current and Previous Version", "RAT File", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message + "\n" + ex.Source, "RAT", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        private void txtCurrVer_TextChanged(object sender, TextChangedEventArgs e)
        {
            string path = Path.GetFileName(this.txtCurrVer.Text) + "_RAT.csv";
            this.txtRATFile.Text = Path.Combine(this.txtCurrVer.Text, path);
        }

        private void removeJobs(FileSystemInfo removeXML, DirectoryInfo PrevDraft, StreamWriter FileWriter)
        {
            String Dir = PrevDraft.ToString();
            String removeFile = removeXML.ToString();
            String location = PrevDraft + removeFile;

            if (File.Exists(location))
            {
                    XDocument xDocument = XDocument.Load(location);
                    IEnumerable<XElement> source =
                        from job in xDocument.Descendants("JOB")
                        select job;
                    List<XElement> nodeOldJobs = source.ToList<XElement>();

                    nodeOldJobs.ForEach(delegate(XElement oldJob)
                    {
                        FileWriter.WriteLine("{0}, Job, Removed, ,", oldJob.Attribute("JOBNAME").Value);
                    });
            }
            
            
        }
    }
    public static class Xsi
    {
        public static XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";
        public static XName schemaLocation = Xsi.xsi + "schemaLocation";
        public static XName noNamespaceSchemaLocation = Xsi.xsi + "noNamespaceSchemaLocation";
    }


}
