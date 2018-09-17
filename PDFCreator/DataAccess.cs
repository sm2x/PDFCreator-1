using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFCreator
{
    public class DataAccess
    {
        public bool hasData;
        Object dbObj;

        public DataAccess()
        {
            hasData = true;
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
    }
}
