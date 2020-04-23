﻿using System;
using System.Collections.Generic;
using System.Text;

namespace Saibaiwei
{
    public class IntoEngish
    {

        private string[] StrNO = new string[19];
        private string[] Unit = new string[8];
        private string[] StrTens = new string[9];

        #region NumberToString
        public string NumberToString(double Number)
        {
            string Str;
            string BeforePoint;
            string AfterPoint;
            string tmpStr;
            int nBit;
            string CurString;
            int nNumLen;
            Init();
            Str = Convert.ToString(Math.Round(Number, 2));
            if (Str.IndexOf(".") == -1)
            {
                BeforePoint = Str;
                AfterPoint = "";
            }
            else
            {
                BeforePoint = Str.Substring(0, Str.IndexOf("."));
                AfterPoint = Str.Substring(Str.IndexOf(".") + 1, Str.Length - Str.IndexOf(".") - 1);
            }
            if (BeforePoint.Length > 12)
            {
                return null;
            }
            Str = "";
            while (BeforePoint.Length > 0)
            {
                nNumLen = Len(BeforePoint);
                if (nNumLen % 3 == 0)
                {
                    CurString = Left(BeforePoint, 3);
                    BeforePoint = Right(BeforePoint, nNumLen - 3);
                }
                else
                {
                    CurString = Left(BeforePoint, (nNumLen % 3));
                    BeforePoint = Right(BeforePoint, nNumLen - (nNumLen % 3));
                }
                nBit = Len(BeforePoint) / 3;
                tmpStr = DecodeHundred(CurString);
                if ((BeforePoint == Len(BeforePoint).ToString() || nBit == 0) && Len(CurString) == 3)
                {
                    if (System.Convert.ToInt32(Left(CurString, 1)) != 0 & System.Convert.ToInt32(Right(CurString, 2)) != 0)
                    {
                        tmpStr = Left(tmpStr, tmpStr.IndexOf(Unit[3]) + Len(Unit[3])) + Unit[7] + " " + Right(tmpStr, Len(tmpStr) - (tmpStr.IndexOf(Unit[3]) + Len(Unit[3])));
                    }
                    else
                    {
                        tmpStr = Unit[7] + " " + tmpStr;
                    }
                }
                if (nBit == 0)
                {
                    Str = Convert.ToString(Str + " " + tmpStr).Trim();
                }
                else
                {
                    Str = Convert.ToString(Str + " " + tmpStr + " " + Unit[nBit - 1]).Trim();
                }
                if (Left(Str, 3) == Unit[7])
                {
                    Str = Convert.ToString(Right(Str, Len(Str) - 3)).Trim();
                }
                if (BeforePoint == Len(BeforePoint).ToString())
                {
                    return "";
                }
            }
            BeforePoint = Str;
            if (Len(AfterPoint) > 0)
            {
                AfterPoint = Unit[5] + " " + DecodeHundred(AfterPoint) + " " + Unit[6];
            }
            else
            {
                AfterPoint = Unit[4];
            }
            return BeforePoint+" "+"DOLLARS" + " " + AfterPoint;

        }

        #endregion

        #region Init
        private void Init()
        {
            if (StrNO[0] != "ONE")
            {
                StrNO[0] = "ONE";
                StrNO[1] = "TWO";
                StrNO[2] = "THREE";
                StrNO[3] = "FOUR";
                StrNO[4] = "FIVE";
                StrNO[5] = "SIX";
                StrNO[6] = "SEVEN";
                StrNO[7] = "EIGHT";
                StrNO[8] = "NINE";
                StrNO[9] = "TEN";
                StrNO[10] = "ELEVEN";
                StrNO[11] = "TWELVE";
                StrNO[12] = "THIRTEEN";
                StrNO[13] = "FOURTEEN";
                StrNO[14] = "FIFTEEN";
                StrNO[15] = "SIXTEEN";
                StrNO[16] = "SEVENTEEN";
                StrNO[17] = "EIGHTEEN";
                StrNO[18] = "NINETEEN";
                StrTens[0] = "TEN";
                StrTens[1] = "TWENTY";
                StrTens[2] = "THIRTY";
                StrTens[3] = "FORTY";
                StrTens[4] = "FIFTY";
                StrTens[5] = "SIXTY";
                StrTens[6] = "SEVENTY";
                StrTens[7] = "EIGHTY";
                StrTens[8] = "NINETY";
                Unit[0] = "THOUSAND";
                Unit[1] = "MILLION";
                Unit[2] = "BILLION";
                Unit[3] = "HUNDRED";
                Unit[4] = "ONLY";
                Unit[5] = "POINT";
                Unit[6] = "CENT";
                Unit[7] = " AND";
            }
        }

        #endregion

        #region DecodeHundred
        private string DecodeHundred(string HundredString)
        {
            int tmp;
            string rtn = "";
            if (Len(HundredString) > 0 && Len(HundredString) <= 3)
            {
                switch (Len(HundredString))
                {
                    case 1:
                        tmp = System.Convert.ToInt32(HundredString);
                        if (tmp != 0)
                        {
                            rtn = StrNO[tmp - 1].ToString();
                        }
                        break;
                    case 2:
                        tmp = System.Convert.ToInt32(HundredString);
                        if (tmp != 0)
                        {
                            if ((tmp < 20))
                            {
                                rtn = StrNO[tmp - 1].ToString();
                            }
                            else
                            {
                                if (System.Convert.ToInt32(Right(HundredString, 1)) == 0)
                                {
                                    rtn = StrTens[Convert.ToInt32(tmp / 10) - 1].ToString();
                                }
                                else
                                {
                                    rtn = Convert.ToString(StrTens[Convert.ToInt32(tmp / 10) - 1] + "-" + StrNO[System.Convert.ToInt32(Right(HundredString, 1)) - 1]);
                                }
                            }
                        }
                        break;
                    case 3:
                        if (System.Convert.ToInt32(Left(HundredString, 1)) != 0)
                        {
                            rtn = Convert.ToString(StrNO[System.Convert.ToInt32(Left(HundredString, 1)) - 1] + " " + Unit[3] + " " + DecodeHundred(Right(HundredString, 2)));
                        }
                        else
                        {
                            rtn = DecodeHundred(Right(HundredString, 2)).ToString();
                        }
                        break;
                    default:
                        break;
                }
            }
            return rtn;
        }

        #endregion

        #region Left
        private string Left(string str, int n)
        {
            return str.Substring(0, n);
        }
        #endregion

        #region Right
        private string Right(string str, int n)
        {
            return str.Substring(str.Length - n, n);
        }
        #endregion

        #region Len
        private int Len(string str)
        {
            return str.Length;
        }
        #endregion
    }

}