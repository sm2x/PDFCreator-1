using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFCreator
{
    public class DataAccess
    {
        public bool _hasData;
        Object dbObj;

        public DataAccess(bool hasData)
        {
            _hasData = hasData;
            //call the db, hold object here in memory
            //dbObj = db.tbl_BPSMDEMergeFields(1);
        }
        public string GetEmployerName()
        {
            return "Stantonbury Ecumenical Partnership (Milton Keynes)";
        }

        public string GetCessationDate()
        {
            return "23 July 2018";
        }
        public string GetPreviousCEWord()
        {            
            return "Our records indicate that your organisation incurred a cessation event previously, and we will be in touch with you about this separately in due course.";
        }
    }
}
