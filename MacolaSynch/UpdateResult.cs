using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MacolaSynch
{
    class UpdateResult
    {

        private string m_sFieldName;
        private string m_sOldValue;
        private string m_sNewValue;

        public UpdateResult(string FieldName, string OldValue, string NewValue)
        {
            m_sFieldName = FieldName;
            m_sOldValue = OldValue;
            m_sNewValue = NewValue;
        }

        public string FieldName
        {
            get
            {
                return m_sFieldName;
            }
            set
            {
                m_sFieldName = value;
            }
        }

        public string OldValue
        {
            get
            {
                return m_sOldValue;
            }
            set
            {
                m_sOldValue = value;
            }
        }

        public string NewValue
        {
            get
            {
                return m_sNewValue;
            }
            set
            {
                m_sNewValue = value;
            }
        }

    }
}
