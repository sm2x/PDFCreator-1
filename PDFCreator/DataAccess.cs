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

        public string GetMergeData(int mergeFieldId)
        {
            return "";
        }

        //these are for testing purposes, it will return a db obj then the data will be accessed via the columns
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
            if(_hasData)
            {
                return "Our records indicate that your organisation incurred a cessation event previously, and we will be in touch with you about this separately in due course.";
            }
            else
            {
                return "";
            }            
        }
        public string GetCessationAmount()
        {
            return "£200,000";
        }

        public class TableCellsDTO
        {
            public string CellText { get; set; }
            public float Width { get; set; }
            public bool? isBold { get; set; }
        }

        public enum MergeField
        {
            Title,
            EstimatedEmployerDebtAt,
            IntroListItemOnePartOne,
            IntroListItemOnePartTwo,
            IntroListItemTwo,
            IntroListItemThree,
            IntroListItemFour,
            IntroListItemFive,
            IntroListItemSix,
            IntroListItemSeven,
            EstimatedEmployerDebtParagraph,
            ComparisonPreviousFigure,
            DoINeedItemOne,
            DoINeedItemTwo,
            DoINeedItemThree,
            HowDebtCalculatedItemOne,
            HowDebtCalculatedItemTwo,
            HowDebtCalculatedItemThree,
            HowDebtCalculatedItemFour,
            HowDebtCalculatedItemFive,
            HowDoesRelateItemOne,
            HowDoesRelateItemTwo,
            HowDoesRelateItemThree
        }


        public string GetText(MergeField field)
        {
            var text = "";
            switch (field)
            {
                case MergeField.Title:
                    text = string.Format("The Baptist Pension Scheme (BPS) – {0}", GetEmployerName());
                    break;
                case MergeField.EstimatedEmployerDebtAt:
                    text = string.Format("Estimated Employer Debt as at {0} ", GetCessationDate());
                    break;
                case MergeField.IntroListItemOnePartOne:
                    text = "You do not need to take any action ";
                    break;
                case MergeField.IntroListItemOnePartTwo:
                    text = "as a result of this document, which is for guidance only.  It provides an estimate of the employer debt that your organisation would need to pay, if it were to exit the defined benefit section of the BPS by paying its employer debt immediately.";
                    break;
                case MergeField.IntroListItemTwo:
                    text = "The BPS and its advisers/administrators accept no liability to any organisation for any actions taken (or not taken) as a result of this estimate, the accompanying guidance notes and FAQs. It is each organisation’s responsibility to ensure that it understands the complex legal position in relation to Employer Debts, taking professional advice as necessary.";
                    break;
                case MergeField.IntroListItemThree:
                    text = "There are a number of reasons why the actual figure in your circumstances could differ significantly from the figure set out below.  Please read the notes carefully in case this applies to you.";
                    break;
                case MergeField.IntroListItemFour:
                    text = "Updated figures will be provided on a monthly basis and will rise and fall over time, depending on how the financial position of the Scheme alters.";
                    break;
                case MergeField.IntroListItemFive:
                    text = "This document is for your information. It is not a demand for payment, and you do not need to take any action:";
                    break;
                case MergeField.IntroListItemSix:
                    text = "If your organisation has incurred and/or settled a debt in the last 3 months then this estimate might not reflect your up-to-date position.  This is because the details of each employer’s status that are used in the estimated debt calculation are updated once per calendar quarter, so will not reflect more recent cessation events or debt payments.";
                    break;
                case MergeField.IntroListItemSeven:
                    text = "The estimate provided in this document might not reflect the total amount that would be due if your organisation incurs a cessation event.  In particular: \n";
                    break;
                case MergeField.EstimatedEmployerDebtParagraph:
                    text = string.Format("The estimated employer debt for your organisation at {0} is {1}.  The calculation is summarised in Table 1 below. This figure has been calculated using this “assumed” cessation date based on financial conditions near the month end, but unless your organisation happened to have an actual cessation event on this date, no employer debt is due to be paid at this time based on this cessation date and this estimated debt.  The figure is for information only.", GetCessationDate(), GetCessationAmount());
                    break;
                case MergeField.ComparisonPreviousFigure:
                    text = "Figures for individual employers may vary from month to month as a result of changes in Scheme membership (for example, retirements, deaths or transfers out of the Scheme), as well as reflecting the general trend.";
                    break;
                case MergeField.DoINeedItemOne:
                    text = "This document is not a demand for payment, and you do not need to take any action.  Your organisation can simply use this information for monitoring the changes to the estimated employer debt over time, while continuing to make the required monthly deficit contributions, provided it continues to employ an active member of the BPS.";
                    break;
                case MergeField.DoINeedItemTwo:
                    text = "If your organisation were to consider incurring a cessation event and settling its employer debt, you should note the following:";
                    break;
                case MergeField.DoINeedItemThree:
                    text = "To understand fully the implications of and timescales for any decision, you must also read the accompanying documents:";
                    break;
                case MergeField.HowDebtCalculatedItemOne:
                    text = "The estimated employer debt is calculated based on the BPS’ funding position and the Trustee's knowledge of the BPS’ liabilities as at the assumed cessation date, based on financial conditions near the month end.  More details are provided in the Guidance Notes and Frequently Asked Questions, and a summary of the calculation for your organisation is shown in Table 1.";
                    break;
                case MergeField.HowDebtCalculatedItemTwo:
                    text = "The estimated employer debt is calculated using the Trustee's record of your organisation’s current and former ministers who were in service with you, as set out in Table 2 of this document. If you do not agree with the record in Table 2, please let us know by emailing Mark Hynes, the Pensions Manager on mhynes@baptist.org.uk";
                    break;
                case MergeField.HowDebtCalculatedItemThree:
                    text = "Your organisation’s pension information is confidential to the BPS and cannot be shared without your permission. However, the Baptist Union Regional Associations have supported a number of churches and other scheme employers understand their obligations and options and you may find it helpful to contact them.";
                    break;
                case MergeField.HowDebtCalculatedItemFour:
                    text = "The size of an employer’s liability to the BPS depends on two main factors:";
                    break;
                case MergeField.HowDebtCalculatedItemFive:
                    text = "It is important to note that the Employer Debt Regulations (2005 and as amended), require any liabilities which cannot specifically be attributed to any current employer (“orphan liabilities”) to be shared amongst all current employers.  These orphan liabilities are part of the Scheme deficit calculation.";
                    break;
                case MergeField.HowDoesRelateItemOne:
                    text = "Your monthly deficiency payments to the defined benefit plan are set every three years to target the deficit in the Scheme at the time.  This deficit is measured using different assumptions from those used to calculate the estimated employer debt.  As a result the monthly contributions are targeting a lower deficit than the very prudent measure that the regulations say must be used for employer debt calculations.";
                    break;
                case MergeField.HowDoesRelateItemTwo:
                    text = "The monthly deficiency payments are at present similar for most employers.  This contrasts with the employer debt calculations, which depend directly on the liabilities in the BPS that relate to each employer.";
                    break;
                case MergeField.HowDoesRelateItemThree:
                    text = "As a result, your estimated employer debt could look very large, or indeed quite small, compared with the amount you might expect to pay on a monthly basis over the period to 2028 (which is the end point of the current plan to address the Scheme deficit).";
                    break;
                default:
                    break;
            }
            return text;
        }
    }
}
