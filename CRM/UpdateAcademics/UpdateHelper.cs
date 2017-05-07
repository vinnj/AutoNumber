using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;
using LinqToExcel.Attributes;
using System.Collections;

namespace UpdateHelper
{
    #region UpdateAcademics
    public class UpdateAcademics
    {
        //DECLARE XL COLUMNS TO READ
        [ExcelColumn("Email")]
        public string Email { get; set; }

        [ExcelColumn("ContactID")]
        public string ContactID { get; set; }
        
        //USE LINQTOEXCEL METHODS TO PARSE THE XL FILE
        public ExcelQueryFactory ReadAcademicData()
        {
         
            var academicExcel = new ExcelQueryFactory(@"C:\TestAcademicUpdate\VinnTest.xlsx")
            {
                DatabaseEngine = LinqToExcel.Domain.DatabaseEngine.Ace,
                TrimSpaces = LinqToExcel.Query.TrimSpacesType.Both,
                UsePersistentConnection = true,
                ReadOnly = true
            };

            return academicExcel;

            }
    
        }

  }

     
      


  

#endregion
