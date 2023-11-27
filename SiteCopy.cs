#region Summary
/******************************************************************************
// AUTHOR                   : Mark Nischalke 
// CREATE DATE              : 4/14/10 
// PURPOSE                  : Add custom command to stsadm
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
using Microsoft.SharePoint.StsAdmin;
using Microsoft.SharePoint;

namespace MANSoftdev.SharePoint.StsadmEx
{
    public class CopySite : ISPStsadmCommand
    {
        public CopySite()
        {

        }

        #region ISPStsadmCommand Members

        public string GetHelpMessage(string command)
        {
            string msg = "unrecognized command";
            if(command.ToLower() == "copy")
            {
                msg = "-o copyweb \r\n -source <url of source web> -dest <url of destination web>";
            }
            else if(command.ToLower() == "move")
            {
                msg = "-o moveweb \r\n -source <url of source web> -dest <url of destination web>";
            }

            return msg;

        }

        public int Run(string command, System.Collections.Specialized.StringDictionary keyValues, out string output)
        {
            if(!keyValues.ContainsKey("source"))
            {
                throw new InvalidOperationException("Source url must be specified");
            }

            if(!keyValues.ContainsKey("dest"))
            {
                throw new InvalidOperationException("Destination url must be specified");
            }

            MANSoftdev.SharePoint.StsadmEx.SPWeb web = new SPWeb();
            web.SourceURL = keyValues["source"];
            web.DestinationURL = keyValues["destination"];

            if(command.ToLower() == "copy")
            {
                web.Copy();   
            }
            else if(command.ToLower() == "move")
            {
                web.Move();
            }

            output = "Success";
            return 0;
        }

        #endregion
    }
}
