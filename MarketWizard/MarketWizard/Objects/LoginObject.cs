using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarketWizard
{
    public class LoginObject
    {
        private string emailAddress;
        private string pwd;

        public LoginObject()
        {
            emailAddress = "NULL";
            pwd = "NULL";
        }

        public void setAddress(string address)
        {
            emailAddress = address;
        }

        public void setPassword(string password)
        {
            pwd = password;
        }

        public string getAddress()
        {
            return emailAddress;
        }

        public string getPassword()
        {
            return pwd;
        }
    }
}
