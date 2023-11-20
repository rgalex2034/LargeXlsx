using System;
using System.Collections.Generic;
using System.Text;

namespace LargeXlsx
{
    public class XlsxProtection
    {
        public static XlsxProtection Default = new XlsxProtection(locked: true, hiddenFormula: false);
        public bool Locked { get; }
        public bool HiddenFormula { get; }

        public XlsxProtection(bool locked, bool hiddenFormula)
        {
            Locked = locked;
            HiddenFormula = hiddenFormula;
        }

        public XlsxProtection WithLocked(bool locked) => new XlsxProtection(locked, HiddenFormula);
        public XlsxProtection WithHiddenFormula(bool hiddenFormula) => new XlsxProtection(Locked, hiddenFormula);

        public override bool Equals(object obj)
        {
            return Equals(obj as XlsxProtection);
        }

        public bool Equals(XlsxProtection other)
        {
            return other != null && Locked == other.Locked && HiddenFormula == other.HiddenFormula;
        }

        public override int GetHashCode()
        {
            var hashCode = 493172489;
            hashCode = hashCode * -1521134295 + Locked.GetHashCode();
            hashCode = hashCode * -1521134295 + HiddenFormula.GetHashCode();
            return hashCode;
        }

        public static bool operator ==(XlsxProtection protection1, XlsxProtection protection2)
        {
            return EqualityComparer<XlsxProtection>.Default.Equals(protection1, protection2);
        }

        public static bool operator !=(XlsxProtection protection1, XlsxProtection protection2)
        {
            return !(protection1 == protection2);
        }
    }
}
