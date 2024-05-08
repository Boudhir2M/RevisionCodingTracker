using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Révision;

namespace Révision
{
    public static class Calcul
    {
        public static string ChiffreToLettre(string Num)
        {
            string tempChiffreToLettre = null;
            string snum;
            snum = Convert.ToString(Num);
            string centim;
            centim = snum.Length >= 2 ? snum.Substring(snum.Length - 2) : snum;
            string decimals = null;
            int position = 0;
            int l = 0;
            string[] T = null;
            int r = 0;
            int p = 0;
            l = snum.Length;
            snum = snum.Replace(".", ",");
            position = snum.IndexOf(",") + 1;
            if (position == 0)
            {
                centim = "00";
                decimals = snum;
            }
            else
            {
                centim = snum.Substring(position, l - position);
                if (centim.Length == 1)
                {
                    centim = centim + "0";
                }
                decimals = snum.Substring(0, position - 1);
            }

            l = decimals.Length;
            r = l % 3;
            p = l / 3;
            if (r > 0)
            {
                p = p + 1;
            }
            T = new string[p + 1];
            for (var i = 0; i < p; i++)
            {
                if (i == 0)
                {
                    if (r > 0)
                    {
                        T[i] = Num.Substring(0, r);
                    }
                    else
                    {
                        T[i] = Num.Substring(0, 3);
                    }
                }
                else
                {
                    if (r > 0)
                    {
                        T[i] = Num.Substring((r + 1 + 3 * (i - 1)) - 1, 3);
                    }
                    else
                    {
                        T[i] = Num.Substring((1 + 3 * i) - 1, 3);
                    }
                }
            }
            string str = "";
            string s = null;

            for (var i = p - 1; i >= 0; i--)
            {
                switch (i)
                {
                    case 0:
                        s = ChiffreToLettre3(T[p - 1 - i]);
                        if (!string.IsNullOrEmpty(s))
                        {
                            str = str + s;
                        }
                        break;
                    case 1:
                        s = ChiffreToLettre3(T[p - 1 - i]);
                        if (!string.IsNullOrEmpty(s))
                        {
                            str = str + s + " mille ";
                        }
                        break;
                    case 2:
                        s = ChiffreToLettre3(T[p - 1 - i]);
                        if (!string.IsNullOrEmpty(s))
                        {
                            str = str + s + " million ";
                        }
                        break;
                    case 3:
                        s = ChiffreToLettre3(T[p - 1 - i]);
                        if (!string.IsNullOrEmpty(s))
                        {
                            str = str + s + " millard ";
                        }
                        break;
                    case 4:
                        s = ChiffreToLettre3(T[p - 1 - i]);
                        if (!string.IsNullOrEmpty(s))
                        {
                            str = str + s + " billion ";
                        }
                        break;
                    case 5:
                        s = ChiffreToLettre3(T[p - 1 - i]);
                        if (!string.IsNullOrEmpty(s))
                        {
                            str = str + s + " billiard ";
                        }
                        break;
                    case 6:
                        s = ChiffreToLettre3(T[p - 1 - i]);
                        if (!string.IsNullOrEmpty(s))
                        {
                            str = str + s + " trillion ";
                        }
                        break;
                }
            }

            if (centim == "00")
            {
                tempChiffreToLettre = str + ".";
            }
            else
            {
                tempChiffreToLettre = str + " et " + ChiffreToLettre3(centim) + " centimes.";
            }

            return tempChiffreToLettre;
        }

        public static string ChiffreToLettre3(string Num)
        {
            long l = 0;
            l = Num.Length;
            string str = "";
            string cent = "";
            string dix = "";
            string unite = "";
            switch (l)
            {
                case 0:
                    cent = "";
                    dix = "";
                    unite = "";
                    break;
                case 1:
                    cent = "";
                    dix = "";
                    unite = Num;
                    break;
                case 2:
                    cent = "";
                    dix = Num.Substring(0, 1);
                    unite = Num.Substring(Num.Length - 1, 1);
                    break;
                case 3:
                    cent = Num.Substring(0, 1);
                    dix = Num.Substring(1, 1);
                    unite = Num.Substring(Num.Length - 1, 1);
                    break;
            }

            switch (cent)
            {
                case "0":
                    break;
                case "1":
                    str = "cent";
                    break;
                case "2":
                    str = "deux cent";
                    break;
                case "3":
                    str = "trois cent";
                    break;
                case "4":
                    str = "quatre cent";
                    break;
                case "5":
                    str = "cinq cent";
                    break;
                case "6":
                    str = "six cent";
                    break;
                case "7":
                    str = "sept cent";
                    break;
                case "8":
                    str = "huit cent";
                    break;
                case "9":
                    str = "neuf cent";
                    break;
            }

            switch (dix)
            {
                case "0":
                    break;
                case "1":
                    switch (unite)
                    {
                        case "0":
                            str = str + " dix";
                            break;
                        case "1":
                            str = str + " onze";
                            break;
                        case "2":
                            str = str + " douze";
                            break;
                        case "3":
                            str = str + " treize";
                            break;
                        case "4":
                            str = str + " quatorze";
                            break;
                        case "5":
                            str = str + " quinze";
                            break;
                        case "6":
                            str = str + " seize";
                            break;
                        case "7":
                            str = str + " dix sept";
                            break;
                        case "8":
                            str = str + " dix huit";
                            break;
                        case "9":
                            str = str + " dix neuf";
                            break;
                    }
                    break;
                case "2":
                    str = str + " vingt";
                    break;
                case "3":
                    str = str + " trente";
                    break;
                case "4":
                    str = str + " quarante";
                    break;
                case "5":
                    str = str + " cinquante";
                    break;
                case "6":
                    str = str + " soixante";
                    break;
                case "7":
                    switch (unite)
                    {
                        case "0":
                            str = str + " soixante dix";
                            break;
                        case "1":
                            str = str + " soixante onze";
                            break;
                        case "2":
                            str = str + " soixante douze";
                            break;
                        case "3":
                            str = str + " soixante treize";
                            break;
                        case "4":
                            str = str + " soixante quatorze";
                            break;
                        case "5":
                            str = str + " soixante quinze";
                            break;
                        case "6":
                            str = str + " soixante seize";
                            break;
                        case "7":
                            str = str + " soixante dix sept";
                            break;
                        case "8":
                            str = str + " soixante dix huit";
                            break;
                        case "9":
                            str = str + " soixante dix neuf";
                            break;
                    }
                    break;
                case "8":
                    str = str + " quatre vingt";
                    break;
                case "9":
                    switch (unite)
                    {
                        case "0":
                            str = str + " quatre vingt dix";
                            break;
                        case "1":
                            str = str + " quatre vingt onze";
                            break;
                        case "2":
                            str = str + " quatre vingt douze";
                            break;
                        case "3":
                            str = str + " quatre vingt treize";
                            break;
                        case "4":
                            str = str + " quatre vingt quatorze";
                            break;
                        case "5":
                            str = str + " quatre vingt quinze";
                            break;
                        case "6":
                            str = str + " quatre vingt seize";
                            break;
                        case "7":
                            str = str + " quatre vingt dix sept";
                            break;
                        case "8":
                            str = str + " quatre vingt dix huit";
                            break;
                        case "9":
                            str = str + " quatre vingt dix neuf";
                            break;
                    }
                    break;
            }

            if (dix == "1" || dix == "7" || dix == "9")
            {
                // ...
            }
            else
            {
                switch (unite)
                {
                    case "0":
                        break;
                    case "1":
                        str = str + " un";
                        break;
                    case "2":
                        str = str + " deux";
                        break;
                    case "3":
                        str = str + " trois";
                        break;
                    case "4":
                        str = str + " quatre";
                        break;
                    case "5":
                        str = str + " cinq";
                        break;
                    case "6":
                        str = str + " six";
                        break;
                    case "7":
                        str = str + " sept";
                        break;
                    case "8":
                        str = str + " huit";
                        break;
                    case "9":
                        str = str + " neuf";
                        break;
                }
            }
            return str;
        }
    }
}
