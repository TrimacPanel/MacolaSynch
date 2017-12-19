using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MacolaSynch
{
    class AlertItem
    {


        private string m_sItemNo;
        private string m_sDescription;
        private bool m_bActionNeeded;
        private Nullable<double> m_dblMacolaQOH = null;
        private Nullable<double> m_dblAccessQOH = null;

        private AlertTypeEnum m_iType;
        private AlertSeverityEnum m_iSeverity;

        public enum AlertSeverityEnum
        {
            Information,
            Moderate,
            Severe
        }

        public enum AlertTypeEnum
        {
            Add,
            Delete,
            Variance
        }

        public AlertItem(string ItemNo, string Description, AlertTypeEnum AlertType, AlertSeverityEnum Severity, bool ActionNeeded, Nullable<double> MacolaQOH, Nullable<double> AccessQOH)
        {
            m_sItemNo = ItemNo;
            m_sDescription = Description;
            m_iType = AlertType;
            m_iSeverity = Severity;
            m_bActionNeeded = ActionNeeded;
            m_dblMacolaQOH = MacolaQOH;
            m_dblAccessQOH = AccessQOH;
        }
        
        public string ItemNo
        {
            get
            {
                return m_sItemNo;
            }
            set
            {
                m_sItemNo = value;
            }
        }

        public string Description
        {
            get
            {
                return m_sDescription;
            }
            set
            {
                m_sDescription = value;
            }
        }

        public Nullable<double> MacolaQOH
        {
            get
            {
                return m_dblMacolaQOH;
            }
            set
            {
                m_dblMacolaQOH = value;
            }
        }

        public Nullable<double> AccessQOH
        {
            get
            {
                return m_dblAccessQOH;
            }
            set
            {
                m_dblAccessQOH = value;
            }
        }

        public AlertTypeEnum Type 
        {
            get
            {
                return m_iType;
            }
            set
            {
                m_iType = value;
            }
        }


        public AlertSeverityEnum Severity 
        {
            get
            {
                return m_iSeverity;
            }
            set
            {
                m_iSeverity = value;
            }
        }

        public bool ActionNeeded
        {
            get
            {
                return m_bActionNeeded;
            }
            set
            {
                m_bActionNeeded = value;
            }
        }


    }
}
