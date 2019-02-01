/*
Created by: John Walker
Last modified: 1/31/2018
Description: This class is for XML credential save type
*/


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AdminRDP
{
    public class Credential_Manager_Save
    {
        private string txtusername;
        private string txtpassword;
        private string txtdomain;

        public string txtUsername
        {
            get { return txtusername; }
            set { txtusername = value; }
        }

        public string txtPassword
        {
            get { return txtpassword; }
            set { txtpassword = value; }
        }

        public string txtDomain
        {
            get { return txtdomain; }
            set { txtdomain = value; }
        }
    }
}
