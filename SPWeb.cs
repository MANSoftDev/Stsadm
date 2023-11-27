#region Summary
/******************************************************************************
// AUTHOR                   : Mark Nischalke 
// CREATE DATE              : 4/11/2010 
//
// Copyright © MANSoftDev 2010 all rights reserved
// ===========================================================================
// Copyright notice must remain
//
******************************************************************************/
#endregion

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Deployment;
using System.IO;
using Microsoft.SharePoint;
using System.Data.SqlClient;
using Microsoft.SharePoint.Navigation;

namespace MANSoftdev.SharePoint.StsadmEx
{
    public class SPWeb
    {
        private const string EXPORT_FILENAME = "export.cmp";
        private string m_ExportPath;

        #region Constructor/Finalizer

        /// <summary>
        /// Constructor
        /// </summary>
        public SPWeb()
        {

        }

        /// <summary>
        /// Constructor
        /// </summary>
        public SPWeb(string sourceURL, string destinationURL)
        {
            SourceURL = sourceURL;
            DestinationURL = destinationURL;
        }

        /// <summary>
        /// Finalizer
        /// </summary>
        ~SPWeb()
        {
            try
            {
                if(SourceSite != null)
                    SourceSite.Dispose();

                if(DestinationSite != null)
                    DestinationSite.Dispose();

                if(SourceWeb != null)
                    SourceWeb.Dispose();

                if(DestinationWeb != null)
                    DestinationWeb.Dispose();
            }
            catch(InvalidOperationException)
            {
                // Handle is invalid?     
            }

            string path = Path.Combine(ExportPath, EXPORT_FILENAME);
            if(File.Exists(path))
                File.Delete(path);
        }

        #endregion

        #region Copy Methods

        /// <summary>
        /// Copy web from soureURL to destinationURL
        /// Will leave source web in place
        /// </summary>
        public void Copy()
        {
            Run(true);
        }

        /// <summary>
        /// Copy web from soureURL to destinationURL
        /// Will leave source web in place
        /// </summary>
        /// <param name="sourceURL">URL of web to copy</param>
        /// <param name="destinationURL">URL to copy web to</param>
        public void Copy(string sourceURL, string destinationURL)
        {
            SourceURL = sourceURL;
            DestinationURL = destinationURL;
            Run(true);
        }

        #endregion

        #region Move Methods

        /// <summary>
        /// Move web from soureURL to destinationURL
        /// Will remove source web after completion
        /// </summary>
        public void Move()
        {
            Run(false);
        }

        /// <summary>
        /// Move web from soureURL to destinationURL
        /// Will remove source web after completion
        /// </summary>
        /// <param name="sourceURL">URL of web to move</param>
        /// <param name="destinationURL">URL to move web to</param>
        public void Move(string sourceURL, string destinationURL)
        {
            SourceURL = sourceURL;
            DestinationURL = destinationURL;
            Run(false);
        }

        #endregion

        #region Private Methods

        private void Run(bool isCopy)
        {
            // First validate the source and set the set and web
            ValidateSource();

            // Validate the destination site and web
            ValidateDestination();

            // Export the source web
            Export();

            // Import to the destination
            Import();

            // If moving, delete the source web
            if(!isCopy)
            {
                RecursivelyDeleteWeb(SourceWeb);
            }

            // TODO: Add to top link bar and quick launch if necessay
        }

        /// <summary>
        /// Validate the source site and web exists
        /// </summary>
        private void ValidateSource()
        {
            try
            {
                Uri uri = new Uri(SourceURL);

                SourceSite = new Microsoft.SharePoint.SPSite(SourceURL);
                if(SourceSite != null)
                {
                    SourceWeb = SourceSite.OpenWeb(uri.LocalPath);
                    if(!SourceWeb.Exists)
                    {
                        SourceWeb.Dispose();
                        SourceWeb = null;
                        throw new ArgumentException("Source web is invalid");
                    }
                }
            }
            catch(System.IO.FileNotFoundException)
            {
                // Could find the specified site
                throw new ArgumentException("Source site is invalid");
            }
        }

        /// <summary>
        /// Validate destination site and web and create if necessary
        /// </summary>
        private void ValidateDestination()
        {
            try
            {
                // The last segment of the source needs to be
                // appended to the destination to check if it exists
                Uri sourceURI = new Uri(SourceURL);
                Uri uri = new Uri(DestinationURL + "/" + sourceURI.Segments.Last());

                DestinationSite = new Microsoft.SharePoint.SPSite(DestinationURL);
                if(DestinationSite != null)
                {
                    DestinationWeb = DestinationSite.OpenWeb(uri.LocalPath);
                    if(DestinationWeb.Exists)
                    {
                        CompareTemplates();
                    }
                    else
                    {
                        DestinationWeb = null;
                    }
                }
            }
            catch(System.IO.FileNotFoundException)
            {
                // Could find the specified site
                throw new ArgumentException("Destination site is invalid");
            }
        }

        /// <summary>
        /// Compare source and destination templates
        /// </summary>
        private void CompareTemplates()
        {
            uint localID = Convert.ToUInt16(SourceWeb.Locale.LCID);

            string templateName = GetTemplateName(SourceSite.ContentDatabase.DatabaseConnectionString,
                SourceWeb.ID, SourceWeb.WebTemplate);

            SPWebTemplate sourceTemplate = SourceSite.GetWebTemplates(localID)[templateName];

            templateName = GetTemplateName(DestinationSite.ContentDatabase.DatabaseConnectionString,
                DestinationWeb.ID, DestinationWeb.WebTemplate);

            SPWebTemplate destTemplate = DestinationWeb.Site.GetWebTemplates(localID)[templateName];

            // If template are the same then the 
            // destination must be deleted.
            if(sourceTemplate.Name != destTemplate.Name)
            {
                RecursivelyDeleteWeb(DestinationWeb);
                DestinationWeb.Dispose();
                DestinationWeb = null;
            }
        }

        /// <summary>
        /// Get the WebTemplate name for the site matching the given id
        /// </summary>
        /// <param name="connString">ConnectionString for database</param>
        /// <param name="id">Id of site to lookup</param>
        /// <param name="webTemplate">Name of webtemplate used for site</param>
        /// <returns></returns>
        private string GetTemplateName(string connString, Guid id, string webTemplate)
        {
            string cmdText = string.Format("SELECT ProvisionConfig FROM dbo.Webs WHERE Id = '{0}'", id.ToString());
            int provisionConfig = 0;
            using(SqlConnection conn = new SqlConnection(connString))
            {
                using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                {
                    conn.Open();
                    provisionConfig = Convert.ToInt32(cmd.ExecuteScalar());
                }
            }

            return string.Format("{0}#{1}", webTemplate, provisionConfig);
        }

        /// <summary>
        /// Export the web
        /// </summary>
        private void Export()
        {
            if(SourceSite == null)
                throw new ApplicationException("SourceSite is null");

            SPExportSettings settings = new SPExportSettings();
            settings.FileLocation = ExportPath;
            settings.BaseFileName = EXPORT_FILENAME;
            settings.SiteUrl = SourceSite.Url;

            settings.ExportMethod = SPExportMethodType.ExportAll;
            settings.FileCompression = true;
            settings.IncludeVersions = SPIncludeVersions.All;
            settings.IncludeSecurity = SPIncludeSecurity.All;
            settings.ExcludeDependencies = false;
            settings.ExportFrontEndFileStreams = true;
            settings.OverwriteExistingDataFile = true;

            // Only interested in the SourceWeb
            SPExportObject expObj = new SPExportObject();
            expObj.IncludeDescendants = SPIncludeDescendants.All;
            expObj.Id = SourceWeb.ID;
            expObj.Type = SPDeploymentObjectType.Web;
            settings.ExportObjects.Add(expObj);

            SPExport export = new SPExport(settings);
            export.Run();
        }

        /// <summary>
        /// Import web
        /// </summary>
        private void Import()
        {
            if(DestinationSite == null)
                throw new ApplicationException("DestinationSite is null");

            SPImportSettings settings = new SPImportSettings();

            settings.FileLocation = ExportPath;
            settings.BaseFileName = EXPORT_FILENAME;
            settings.IncludeSecurity = SPIncludeSecurity.All;
            settings.UpdateVersions = SPUpdateVersions.Overwrite;
            settings.RetainObjectIdentity = false;
            settings.SiteUrl = DestinationSite.Url;
            settings.WebUrl = DestinationURL;

            SPImport import = new SPImport(settings);
            import.Run();

            //TODO: Option to rename web
            //DestinationWeb.Name = "";
            //DestinationWeb.Title = "";
        }

        /// <summary>
        /// Delete the give web an all subwebs it contains
        /// </summary>
        /// <param name="web">SPWeb to delete</param>
        private void RecursivelyDeleteWeb(Microsoft.SharePoint.SPWeb web)
        {
            foreach(Microsoft.SharePoint.SPWeb subWeb in web.Webs)
            {
                RecursivelyDeleteWeb(subWeb);
            }

            web.Delete();

            // TODO: Remove from parents TopNavigation and QUickLaunch if move
        }

        #endregion

        #region Properties

        public string SourceURL { get; set; }
        public string DestinationURL { get; set; }

        private Microsoft.SharePoint.SPSite SourceSite { get; set; }
        private Microsoft.SharePoint.SPWeb SourceWeb { get; set; }
        private Microsoft.SharePoint.SPSite DestinationSite { get; set; }
        private Microsoft.SharePoint.SPWeb DestinationWeb { get; set; }

        private string ExportPath
        {
            get
            {
                if(string.IsNullOrEmpty(m_ExportPath))
                    m_ExportPath = Path.GetTempPath();

                return m_ExportPath;
            }
        }

        #endregion
    }
}