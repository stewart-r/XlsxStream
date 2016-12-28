using System.Collections.Generic;
using System.Linq;

namespace XlsxStream
{
    public abstract class SingleAttributeSettingsElement
    {
        public virtual IEnumerable<KeyValuePair<string, string>> ToAttributeKvps()
        {
            return GetType().
                GetProperties().
                Select(p => new KeyValuePair<string, string>(p.Name, p.GetValue(this).ToString()));
        }
    }
}