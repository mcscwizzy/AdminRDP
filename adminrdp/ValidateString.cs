/*
Created by: John Walker
Last modified: 1/31/2018
Description: Validates server names to ensure there are no special characters 
*/


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AdminRDP
{
    class ValidateString
    {
        //validates XML strings
        //if string is valid it will return true
        //if string is not valid it will return false
        public static bool IsValid(string String)
        {
            string specialChars = "!@#$%^&*()[{]};: ,<>/?\'\"\t";

            foreach (var Char in specialChars)
            {
                if (String.Contains(Char))
                {
                    return false;
                }
            }

            string specialNumbers = "01234546789";

            foreach (var Char in specialNumbers)
            {
                if (String.StartsWith(Char.ToString()))
                {
                    return false;
                }
            }

            return true;
        }

        public static bool isValidServerName(string String)
        {
            string specialChars = "!@#$%^&*()[{]};: ,<>/?\'\"\t";

            foreach (var Char in specialChars)
            {
                if (String.Contains(Char))
                {
                    return false;
                }
            }

            return true;
        }

    }
}
