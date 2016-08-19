using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
//using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace ColorMe
{
    public partial class Form1 : Form
    {
        //--------------------------------------------------------------
        public Form1()
        {
            InitializeComponent();
        }
        //--------------------------------------------------------------
        private void Pl_FileLoad_DragEnter(object sender, DragEventArgs e)
        {
            //コントロール内にドラッグされたとき実行される
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                //ドラッグされたデータ形式を調べ、ファイルのときはコピーとする
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                //ファイル以外は受け付けない
                e.Effect = DragDropEffects.None;
            }
        }
        //--------------------------------------------------------------
        private void Pl_FileLoad_DragDrop(object sender, DragEventArgs e)
        {
            //ドロップされたすべてのファイル名を取得する
            string[] FileNameList = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            for (int i = 0; i < FileNameList.Length; i++)
            {
                if (ExecConvertCsv(FileNameList[i]) == false)
                {
                    MessageBox.Show("読み込みエラー");
                    return;
                }
            }
        }
        //--------------------------------------------------------------
        // コンバート処理実行
        public bool ExecConvertCsv(string strFileName)
        {
            // ヘッダー
            string[] strHeader;
            // 読み込みCSVファイル種別
            EnumReadCsvClass enReadCsvClassFlg;
            // データー
            List<string[]> listItemData = new List<string[]>();
            // 納品書用
            List<string[]> listReceiptItemData = new List<string[]>();

            // ファイルを読み込む
            StreamReader clsSr;
            try
            {
                clsSr = new StreamReader(strFileName, System.Text.Encoding.GetEncoding("shift_jis"));
                try
                {
                    string strLine;

                    // 1行目読み込み
                    if ((strLine = clsSr.ReadLine()) == null)
                    {
                        throw new FormatException();
                    }
                    // ヘッダーの文字列を読み、種類を判断する
                    strHeader = strLine.Split(',');
                    if (strHeader[0] == CDefine.READ_CSV_CLASS_CONF_STR[(int)EnumReadCsvClass.COLOR_ME])
                    {
                        enReadCsvClassFlg = EnumReadCsvClass.COLOR_ME;
                    }
                    else
                    {
                        throw new FormatException();
                    }
                    // データーを読み込む                
                    while ((strLine = clsSr.ReadLine()) != null)
                    {
                        if (strLine == "")
                        {
                            break;
                        }
                        if (enReadCsvClassFlg == EnumReadCsvClass.COLOR_ME)
                        {
                            // カラーミーの場合は改行が含まれている可能性がある
                            string strAddLine = "";
                            do
                            {
                                strLine += strAddLine;
                                if ((strLine.Length - strLine.Replace("\"", "").Length) % 2 == 0)
                                {
                                    break;
                                }
                            } while ((strAddLine = clsSr.ReadLine()) != null);
                        }
                        // 全てのデーターを格納
                        listItemData.Add(strLine.Split(','));
                        listReceiptItemData.Add(strLine.Split(','));
                    }
                }
                catch
                {
                    return false;
                }
                finally
                {
                    clsSr.Close();
                }
            }
            catch
            {
                return false;
            }

            // コンバート・書き込み処理
            switch (enReadCsvClassFlg)
            {
                case EnumReadCsvClass.COLOR_ME:
                    // 送り状CSV作成
                    if (ConvertCsv(strFileName, enReadCsvClassFlg, listItemData) == false)
                    {
                        return false;
                    }
                    // 納品書作成
                    if (ReceiptConvertCsv(strFileName, enReadCsvClassFlg, listReceiptItemData) == false)
                    {
                        return false;
                    }
                    break;
            }

            return true;
        }
        //--------------------------------------------------------------
        // コンバート処理実行
        public bool ConvertCsv(string strFileName, EnumReadCsvClass enReadCsvClassFlg, List<string[]> listItemData)
        {
            // ゆうプリR出力ファイル名
            string YpprFileName = Path.GetDirectoryName(strFileName) + "\\" + "ゆうプリR.csv";
            // ゆうプリRデーター
            string[] strYpprData = new string[(int)EnumYpprItem.MAX];

            // ヤマト出力ファイル名
            string YamatoFileName = Path.GetDirectoryName(strFileName) + "\\" + "ヤマト.csv";
            // ゆうプリRデーター
            string[] strYamatoData = new string[(int)EnumYamatoItem.MAX];

            int iIDPos = (int)EnumColorMeItem.ORDER_ID;


            while (listItemData.Count > 0)
            {
                string[] strItem = new string[listItemData[0].Length];

                // 両サイドの"を取り外す
                for (int i = 0; i < listItemData[0].Length; i++)
                {
                    strItem[i] = listItemData[0][i].Substring(1, listItemData[0][i].Length - 2);
                }

                if (strItem[(int)EnumColorMeItem.DELIVERY_METHOD] == "宅配便")
                {
                    // ヤマト
                    if (LoadYamatoData(strItem, strYamatoData) == false)
                    {
                        return false;
                    }
                    if (WriteYamatoData(YamatoFileName, strYamatoData) == false)
                    {
                        return false;
                    }
                }
                else
                {
                    // ゆうプリR
                    if (LoadYpprData(strItem, strYpprData) == false)
                    {
                        return false;
                    }
                    if (WriteYpprData(YpprFileName, strYpprData) == false)
                    {
                        return false;
                    }
                }
                // データ削除
                string strID = listItemData[0][iIDPos];
                for (int i = listItemData.Count - 1; i >= 0; i--)
                {
                    if (strID == listItemData[i][iIDPos])
                    {
                        listItemData.RemoveAt(i);
                    }
                }

            }

            return true;
        }
        //--------------------------------------------------------------
        /// <summary>
        /// ヤマト読み込み
        /// </summary>
        public bool LoadYamatoData(string[] strItem, string[] strYamatoData)
        {

            // お客様側管理番号
            strYamatoData[(int)EnumYamatoItem.USER_ID] = "1" + strItem[(int)EnumColorMeItem.ORDER_ID];
            // 送り状種別
            if (strItem[(int)EnumColorMeItem.PAYMENT_METHOD] == "商品代引き")
            {
                strYamatoData[(int)EnumYamatoItem.INVOICE_CLASS] = "2";
            }
            else
            {
                strYamatoData[(int)EnumYamatoItem.INVOICE_CLASS] = "0";
            }
            // 発送予定日
            strYamatoData[(int)EnumYamatoItem.SHIPPING_SCHEDULE_TIME] = DateTime.Now.ToString("yyyy/MM/dd");
            // お届け日
            if (strItem[(int)EnumColorMeItem.DELIBERY_DATE] != "設定なし")
            {
                DateTime timeDelibery;
                DateTime timeNow;

                timeDelibery = new DateTime(int.Parse(strItem[(int)EnumColorMeItem.DELIBERY_DATE].Substring(0, 4)),
                                            int.Parse(strItem[(int)EnumColorMeItem.DELIBERY_DATE].Substring(5, 2)),
                                            int.Parse(strItem[(int)EnumColorMeItem.DELIBERY_DATE].Substring(8, 2)),
                                            0, 0, 0
                                            );
                timeNow = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);

                if (timeDelibery.CompareTo(timeNow) > 0)
                {
                    strYamatoData[(int)EnumYamatoItem.DELIBERY_DATE] = strItem[(int)EnumColorMeItem.DELIBERY_DATE].Substring(0, 4) +
                                                                   "/" +
                                                                   strItem[(int)EnumColorMeItem.DELIBERY_DATE].Substring(5, 2) +
                                                                   "/" +
                                                                   strItem[(int)EnumColorMeItem.DELIBERY_DATE].Substring(8, 2);
                }
                else
                {
                    strYamatoData[(int)EnumYamatoItem.DELIBERY_DATE] = "";
                }
            }
            else
            {
                strYamatoData[(int)EnumYamatoItem.DELIBERY_DATE] = "";
            }
            // お届け時間
            if (strItem[(int)EnumColorMeItem.DELIBERY_TIME] == "午前中")
            {
                strYamatoData[(int)EnumYamatoItem.DELIBERY_TIME] = "0812";
            }
            else if (strItem[(int)EnumColorMeItem.DELIBERY_TIME] == "12時頃～14時頃")
            {
                strYamatoData[(int)EnumYamatoItem.DELIBERY_TIME] = "1214";
            }
            else if (strItem[(int)EnumColorMeItem.DELIBERY_TIME] == "14時頃～16時頃")
            {
                strYamatoData[(int)EnumYamatoItem.DELIBERY_TIME] = "1416";
            }
            else if (strItem[(int)EnumColorMeItem.DELIBERY_TIME] == "16時頃～18時頃")
            {
                strYamatoData[(int)EnumYamatoItem.DELIBERY_TIME] = "1618";
            }
            else if (strItem[(int)EnumColorMeItem.DELIBERY_TIME] == "18時頃～20時頃")
            {
                strYamatoData[(int)EnumYamatoItem.DELIBERY_TIME] = "1820";
            }
            else
            {
                strYamatoData[(int)EnumYamatoItem.DELIBERY_TIME] = "";
            }
            // お届け先電話番号
            strYamatoData[(int)EnumYamatoItem.TRANSPORT_TEL] = strItem[(int)EnumColorMeItem.PHONE_NO];
            // お届け先郵便番号
            strYamatoData[(int)EnumYamatoItem.TRANSPORT_POST_NO] = strItem[(int)EnumColorMeItem.POST_NO];
            // お届け先住所
            int iPos;

            iPos = strItem[(int)EnumColorMeItem.ADDRESS_2].IndexOf(" ");
            if (iPos <= 0)
            {
                iPos = strItem[(int)EnumColorMeItem.ADDRESS_2].IndexOf("　");
                if (iPos <= 0)
                {
                    strYamatoData[(int)EnumYamatoItem.TRANSPORT_ADDRESS_1] = strItem[(int)EnumColorMeItem.ADDRESS_1] + strItem[(int)EnumColorMeItem.ADDRESS_2];
                    strYamatoData[(int)EnumYamatoItem.TRANSPORT_ADDRESS_2] = "";
                }
                else
                {
                    strYamatoData[(int)EnumYamatoItem.TRANSPORT_ADDRESS_1] = strItem[(int)EnumColorMeItem.ADDRESS_1] + strItem[(int)EnumColorMeItem.ADDRESS_2].Substring(0, iPos);
                    strYamatoData[(int)EnumYamatoItem.TRANSPORT_ADDRESS_2] = strItem[(int)EnumColorMeItem.ADDRESS_2].Substring(iPos + 1, strItem[(int)EnumColorMeItem.ADDRESS_2].Length - (iPos + 1));
                }
            }
            else
            {
                strYamatoData[(int)EnumYamatoItem.TRANSPORT_ADDRESS_1] = strItem[(int)EnumColorMeItem.ADDRESS_1] + strItem[(int)EnumColorMeItem.ADDRESS_2].Substring(0, iPos);
                strYamatoData[(int)EnumYamatoItem.TRANSPORT_ADDRESS_2] = strItem[(int)EnumColorMeItem.ADDRESS_2].Substring(iPos + 1, strItem[(int)EnumColorMeItem.ADDRESS_2].Length - (iPos + 1));
            }
            // お届け先会社・部門名1
            strYamatoData[(int)EnumYamatoItem.TRANSPORT_COMPANY_1] = "";
            // お届け先会社・部門名2
            strYamatoData[(int)EnumYamatoItem.TRANSPORT_COMPANY_2] = "";
            // お届け先名
            strYamatoData[(int)EnumYamatoItem.TRANSPORT_NAME] = strItem[(int)EnumColorMeItem.NAME];
            // お届け先敬称
            strYamatoData[(int)EnumYamatoItem.TRANSPORT_TITLE] = "様";
            // ご依頼主コード
            strYamatoData[(int)EnumYamatoItem.ORIGIN_CODE] = "";
            // ご依頼主電話番号
            strYamatoData[(int)EnumYamatoItem.ORIGIN_TEL] = "050-3786-7989";
            // ご依頼主郵便番号
            strYamatoData[(int)EnumYamatoItem.ORIGIN_POST_NO] = "4000306";
            // ご依頼主住所1
            strYamatoData[(int)EnumYamatoItem.ORIGIN_ADDRESS_1] = "山梨県南アルプス市小笠原1589-1";
            // ご依頼主住所2
            strYamatoData[(int)EnumYamatoItem.ORIGIN_ADDRESS_2] = "";
            // ご依頼主名
            strYamatoData[(int)EnumYamatoItem.ORIGIN_NAME] = "ASShop";
            // 品名コード1
            strYamatoData[(int)EnumYamatoItem.PRODUCT_NAME_CODE_1] = "";
            // 品名1
            strYamatoData[(int)EnumYamatoItem.PRODUCT_NAME_1] = strItem[(int)EnumColorMeItem.PRODUCT_NAME];
            if (strYamatoData[(int)EnumYamatoItem.PRODUCT_NAME_1].Length > 25)
            {
                strYamatoData[(int)EnumYamatoItem.PRODUCT_NAME_1] = strYamatoData[(int)EnumYamatoItem.PRODUCT_NAME_1].Substring(0, 25);
            }
            // 品名コード2
            strYamatoData[(int)EnumYamatoItem.PRODUCT_NAME_CODE_2] = "";
            // 品名2
            strYamatoData[(int)EnumYamatoItem.PRODUCT_NAME_2] = "";
            // 荷扱い1
            strYamatoData[(int)EnumYamatoItem.FREIGHT_HANDLING_1] = "精密機器";
            // 荷扱い2
            strYamatoData[(int)EnumYamatoItem.FREIGHT_HANDLING_2] = "下積厳禁";
            // 記事
            strYamatoData[(int)EnumYamatoItem.ARTICLE] = "";
            // 代引金額
            if (strItem[(int)EnumColorMeItem.PAYMENT_METHOD] == "商品代引き")
            {
                strYamatoData[(int)EnumYamatoItem.COD_PAY] = strItem[(int)EnumColorMeItem.TOTAL_PAY];
            }
            else
            {
                strYamatoData[(int)EnumYamatoItem.COD_PAY] = "";
            }
            // 代引消費税
            strYamatoData[(int)EnumYamatoItem.COD_TAX] = "";
            // 発行枚数
            strYamatoData[(int)EnumYamatoItem.POST_NUM] = "1";
            // 個数口枠の印字
            strYamatoData[(int)EnumYamatoItem.NUMBER_FRAME] = "3";
            // ご請求先顧客コード
            strYamatoData[(int)EnumYamatoItem.BILLING_CODE] = "05037867989";
            // 運賃管理番号
            strYamatoData[(int)EnumYamatoItem.FARE_NO] = "01";
            // 空白
            strYamatoData[(int)EnumYamatoItem.NULL_1] = "";
            strYamatoData[(int)EnumYamatoItem.NULL_2] = "";
            strYamatoData[(int)EnumYamatoItem.NULL_3] = "";
            strYamatoData[(int)EnumYamatoItem.NULL_4] = "";
            strYamatoData[(int)EnumYamatoItem.NULL_5] = "";
            strYamatoData[(int)EnumYamatoItem.NULL_6] = "";
            strYamatoData[(int)EnumYamatoItem.NULL_7] = "";
            strYamatoData[(int)EnumYamatoItem.NULL_8] = "";
            strYamatoData[(int)EnumYamatoItem.NULL_9] = "";
            strYamatoData[(int)EnumYamatoItem.NULL_10] = "";

            return true;
        }
        //--------------------------------------------------------------
        /// <summary>
        /// ヤマト書き込む
        /// </summary>
        public bool WriteYamatoData(string strFile, string[] strYamatoData)
        {
            bool bHeaderFlg = false;

            // ファイルが存在しない場合には、ヘッダーを追加
            if (System.IO.File.Exists(strFile) == false)
            {
                bHeaderFlg = true;
            }

            StreamWriter clsSw;
            try
            {
                clsSw = new StreamWriter(strFile, true, Encoding.GetEncoding("Shift_JIS"));
            }
            catch
            {
                return false;
            }

            // データ書き込み
            if (bHeaderFlg)
            {
                clsSw.Write("ヤマト運輸データ\n");
            }
            for (int i = 0; i < (int)EnumYamatoItem.MAX; i++)
            {
                clsSw.Write("{0},", strYamatoData[i]);
            }
            clsSw.Write("\n");


            clsSw.Flush();
            clsSw.Close();

            return true;
        }
        //--------------------------------------------------------------
        /// <summary>
        /// ゆうプリR読み込み
        /// </summary>
        public bool LoadYpprData(string[] strItem, string[] strYpprData)
        {

            // お客様側管理番号
            strYpprData[(int)EnumYpprItem.USER_ID] = "";
            // 発送予定日
            strYpprData[(int)EnumYpprItem.SHIPPING_SCHEDULE_DATE] = DateTime.Now.ToString("yyyyMMdd");
            // 発送予定時間
            strYpprData[(int)EnumYpprItem.SHIPPING_SCHEDULE_TIME] = "02";
            // 郵便種別
            if (strItem[(int)EnumColorMeItem.DELIVERY_METHOD] == "宅配便")
            {
                strYpprData[(int)EnumYpprItem.POST_CLASS] = "0";
            }
            else
            {
                strYpprData[(int)EnumYpprItem.POST_CLASS] = "9";
            }
            // 支払元
            if (strItem[(int)EnumColorMeItem.PAYMENT_METHOD] == "商品代引き")
            {
                strYpprData[(int)EnumYpprItem.PAYMENT_SOURCE] = "2";
            }
            else
            {
                strYpprData[(int)EnumYpprItem.PAYMENT_SOURCE] = "0";
            }
            // 送り状種別
            if (strItem[(int)EnumColorMeItem.PAYMENT_METHOD] == "商品代引き")
            {
                strYpprData[(int)EnumYpprItem.INVOICE_CLASS] = "1100783003";
            }
            else if (strItem[(int)EnumColorMeItem.DELIVERY_METHOD] == "宅配便")
            {
                strYpprData[(int)EnumYpprItem.INVOICE_CLASS] = "1100783001";
            }
            else
            {
                strYpprData[(int)EnumYpprItem.INVOICE_CLASS] = "1800800001";
            }
            // お届け先郵便番号
            strYpprData[(int)EnumYpprItem.TRANSPORT_POST_NO] = strItem[(int)EnumColorMeItem.POST_NO];

            // お届け先住所
            strYpprData[(int)EnumYpprItem.TRANSPORT_ADDRESS_1] = strItem[(int)EnumColorMeItem.ADDRESS_1];

            int iPos;

            iPos = strItem[(int)EnumColorMeItem.ADDRESS_2].IndexOf(" ");
            if (iPos <= 0)
            {
                iPos = strItem[(int)EnumColorMeItem.ADDRESS_2].IndexOf("　");
                if (iPos <= 0)
                {
                    strYpprData[(int)EnumYpprItem.TRANSPORT_ADDRESS_2] = strItem[(int)EnumColorMeItem.ADDRESS_2];
                    strYpprData[(int)EnumYpprItem.TRANSPORT_ADDRESS_3] = "";
                }
                else
                {
                    strYpprData[(int)EnumYpprItem.TRANSPORT_ADDRESS_2] = strItem[(int)EnumColorMeItem.ADDRESS_2].Substring(0, iPos);
                    strYpprData[(int)EnumYpprItem.TRANSPORT_ADDRESS_3] = strItem[(int)EnumColorMeItem.ADDRESS_2].Substring(iPos + 1, strItem[(int)EnumColorMeItem.ADDRESS_2].Length - (iPos + 1));
                }
            }
            else
            {
                strYpprData[(int)EnumYpprItem.TRANSPORT_ADDRESS_2] = strItem[(int)EnumColorMeItem.ADDRESS_2].Substring(0, iPos);
                strYpprData[(int)EnumYpprItem.TRANSPORT_ADDRESS_3] = strItem[(int)EnumColorMeItem.ADDRESS_2].Substring(iPos + 1, strItem[(int)EnumColorMeItem.ADDRESS_2].Length - (iPos + 1));
            }

            // お届け先名称1
            strYpprData[(int)EnumYpprItem.TRANSPORT_NAME_1] = strItem[(int)EnumColorMeItem.NAME];
            // お届け先名称2
            strYpprData[(int)EnumYpprItem.TRANSPORT_NAME_2] = "";
            // お届け先敬称
            strYpprData[(int)EnumYpprItem.TRANSPORT_TITLE] = "0";
            // お届け先電話番号
            strYpprData[(int)EnumYpprItem.TRANSPORT_TEL] = strItem[(int)EnumColorMeItem.PHONE_NO];
            // お届け先メール
            strYpprData[(int)EnumYpprItem.TRANSPORT_MAIL] = strItem[(int)EnumColorMeItem.E_MAIL];
            // 発送元郵便番号
            strYpprData[(int)EnumYpprItem.ORIGIN_POST_NO] = "4000306";
            // 発送元住所1
            strYpprData[(int)EnumYpprItem.ORIGIN_ADDRESS_1] = "山梨県";
            // 発送元住所2
            strYpprData[(int)EnumYpprItem.ORIGIN_ADDRESS_2] = "南アルプス市小笠原";
            // 発送元住所3
            strYpprData[(int)EnumYpprItem.ORIGIN_ADDRESS_3] = "1589-1";
            // 発送元名称1
            strYpprData[(int)EnumYpprItem.ORIGIN_NAME_1] = "ASSShop";
            // 発送元名称2
            strYpprData[(int)EnumYpprItem.ORIGIN_NAME_2] = "";
            // 発送元敬称
            strYpprData[(int)EnumYpprItem.ORIGIN_TITLE] = "0";
            // 発送元電話番号
            strYpprData[(int)EnumYpprItem.ORIGIN_TEL] = "05037867989";
            // 発送元メール
            strYpprData[(int)EnumYpprItem.ORIGIN_MAIL] = "shop@assystem.jp";
            // こわれもの
            strYpprData[(int)EnumYpprItem.BREAKABLE_FLG] = "1";
            // 逆さま厳禁
            strYpprData[(int)EnumYpprItem.WAY_UP_FLG] = "1";
            // 下積み厳禁
            strYpprData[(int)EnumYpprItem.DO_NOT_STACK_FLG] = "1";
            // 厚さ
            if (strItem[(int)EnumColorMeItem.DELIVERY_METHOD] == "宅配便")
            {
                strYpprData[(int)EnumYpprItem.THICKNESS] = "";
            }
            else
            {
                strYpprData[(int)EnumYpprItem.THICKNESS] = "10";
            }
            // お届け日
            if (strItem[(int)EnumColorMeItem.DELIBERY_DATE] != "設定なし")
            {
                DateTime timeDelibery;
                DateTime timeNow;

                timeDelibery = new DateTime(int.Parse(strItem[(int)EnumColorMeItem.DELIBERY_DATE].Substring(0, 4)),
                                            int.Parse(strItem[(int)EnumColorMeItem.DELIBERY_DATE].Substring(5, 2)),
                                            int.Parse(strItem[(int)EnumColorMeItem.DELIBERY_DATE].Substring(8, 2)),
                                            0, 0, 0
                                            );
                timeNow = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);

                if (timeDelibery.CompareTo(timeNow) > 0)
                {
                    strYpprData[(int)EnumYpprItem.DELIBERY_DATE] = strItem[(int)EnumColorMeItem.DELIBERY_DATE].Replace("/", "");
                }
                else
                {
                    strYpprData[(int)EnumYpprItem.DELIBERY_DATE] = "";
                }
            }
            else
            {
                strYpprData[(int)EnumYpprItem.DELIBERY_DATE] = "";
            }
            // お届け時間
            if (strItem[(int)EnumColorMeItem.DELIBERY_TIME] == "午前中")
            {
                strYpprData[(int)EnumYpprItem.DELIBERY_TIME] = "51";
            }
            else if (strItem[(int)EnumColorMeItem.DELIBERY_TIME] == "12時頃～14時頃")
            {
                strYpprData[(int)EnumYpprItem.DELIBERY_TIME] = "52";
            }
            else if (strItem[(int)EnumColorMeItem.DELIBERY_TIME] == "14時頃～16時頃")
            {
                strYpprData[(int)EnumYpprItem.DELIBERY_TIME] = "53";
            }
            else if (strItem[(int)EnumColorMeItem.DELIBERY_TIME] == "16時頃～18時頃")
            {
                strYpprData[(int)EnumYpprItem.DELIBERY_TIME] = "54";
            }
            else if (strItem[(int)EnumColorMeItem.DELIBERY_TIME] == "18時頃～20時頃")
            {
                strYpprData[(int)EnumYpprItem.DELIBERY_TIME] = "55";
            }
            else if (strItem[(int)EnumColorMeItem.DELIBERY_TIME] == "20時頃～21時頃")
            {
                strYpprData[(int)EnumYpprItem.DELIBERY_TIME] = "56";
            }
            else
            {
                strYpprData[(int)EnumYpprItem.DELIBERY_TIME] = "00";
            }
            // フリー項目
            strYpprData[(int)EnumYpprItem.FREE_ITEM] = "1" + strItem[(int)EnumColorMeItem.ORDER_ID];
            // 代引金額
            if (strItem[(int)EnumColorMeItem.PAYMENT_METHOD] == "商品代引き")
            {
                strYpprData[(int)EnumYpprItem.COD_PAY] = strItem[(int)EnumColorMeItem.TOTAL_PAY];
            }
            else
            {
                strYpprData[(int)EnumYpprItem.COD_PAY] = "";
            }
            // 代引消費税
            strYpprData[(int)EnumYpprItem.COD_TAX] = "";
            // 商品名設定
            strYpprData[(int)EnumYpprItem.PRODUCT_NAME] = strItem[(int)EnumColorMeItem.PRODUCT_NAME];

            return true;
        }
        //--------------------------------------------------------------
        /// <summary>
        /// ゆうプリR 書き込む
        /// </summary>
        public bool WriteYpprData(string strFile, string[] strYpprData)
        {
            StreamWriter clsSw;
            try
            {
                clsSw = new StreamWriter(strFile, true, Encoding.GetEncoding("Shift_JIS"));
            }
            catch
            {
                return false;
            }

            for (int i = 0; i < (int)EnumYpprItem.MAX; i++)
            {
                clsSw.Write("{0},", strYpprData[i]);
            }
            clsSw.Write("\n");


            clsSw.Flush();
            clsSw.Close();

            return true;
        }
        //--------------------------------------------------------------
        // 納品書コンバート処理実行
        public bool ReceiptConvertCsv(string strFileName, EnumReadCsvClass enReadCsvClassFlg, List<string[]> listReceiptItemData)
        {
            // 商品別データー
            List<string[]> listProductItemData = new List<string[]>();
            int iIDPos = (int)EnumColorMeItem.ORDER_ID;

            //ドキュメントを作成
            const string strPdfFileName = "ReceiptTemp.pdf";
            Document doc = new Document(PageSize.A4, 50, 50, 35, 10);

            try
            {
                //ファイルの出力先を設定
                PdfWriter.GetInstance(doc, new FileStream(strPdfFileName, FileMode.Create));

                //ヘッダーの設定をします。
                HeaderFooter header = new HeaderFooter(new Phrase("納品書・領収書", new Font(BaseFont.CreateFont(@"c:\windows\fonts\msgothic.ttc,1", BaseFont.IDENTITY_H, true), 20)), false);
                //センター寄せ
                header.SetAlignment(ElementTags.ALIGN_CENTER);
                // 線設定
                header.Border = Rectangle.BOX;
                //DocumentにHeaderを設定
                doc.Header = header;

                //フッターの設定をします。
                HeaderFooter footer = new HeaderFooter(new Phrase("〒400-0306 山梨県南アルプス市小笠原1589-1\n株式会社ASsystem  TEL：050-3786-7989\nMAIL：shop@assystem.jp  HP：www.asshop-jp.com", new Font(BaseFont.CreateFont(@"c:\windows\fonts\msgothic.ttc,1", BaseFont.IDENTITY_H, true), 12)), false);
                //センター寄せ
                footer.SetAlignment(ElementTags.ALIGN_LEFT);
                // 線設定
                footer.Border = Rectangle.TOP_BORDER;
                //DocumentにFooterを設定
                doc.Footer = footer;

                //ドキュメントを開く
                doc.Open();

                // 
                while (listReceiptItemData.Count > 0)
                {
                    listProductItemData.Clear();
                    listProductItemData.Add(listReceiptItemData[0]);

                    for (int i = 1; i < listReceiptItemData.Count; i++)
                    {
                        if (listReceiptItemData[i][iIDPos] != listReceiptItemData[0][iIDPos])
                        {
                            continue;
                        }
                        listProductItemData.Add(listReceiptItemData[i]);
                    }

                    // 作成
                    if (MakeColorMeReceipt(doc, listProductItemData) == false)
                    {
                        return false;
                    }

                    // データ削除
                    string strID = listReceiptItemData[0][iIDPos];
                    for (int i = listReceiptItemData.Count - 1; i >= 0; i--)
                    {
                        if (strID == listReceiptItemData[i][iIDPos])
                        {
                            listReceiptItemData.RemoveAt(i);
                        }
                    }

                }
            }
            catch
            {
                return false;
            }
            finally
            {
                //ドキュメントを閉じる
                doc.Close();
            }

            // PDFを開く
            System.Diagnostics.Process.Start(strPdfFileName);

            return true;
        }
        //--------------------------------------------------------------
        /// <summary>
        /// カラーミーの納品書・領収書作成
        /// </summary>
        public bool MakeColorMeReceipt(Document doc, List<string[]> listProductItemData)
        {
            // 両サイドの"を取り外す
            for (int i = 0; i < listProductItemData.Count; i++)
            {
                for (int j = 0; j < listProductItemData[i].Length; j++)
                {
                    listProductItemData[i][j] = listProductItemData[i][j].Substring(1, listProductItemData[i][j].Length - 2);
                }
            }
            // 本文
            Font fnt;
            Paragraph para;
            Table tbl;
            Cell cel;

            fnt = new Font(BaseFont.CreateFont(@"c:\windows\fonts\msgothic.ttc,1", BaseFont.IDENTITY_H, true), 12);
            para = new Paragraph("受注番号：" + listProductItemData[0][(int)EnumColorMeItem.ORDER_ID] + "　　受注日：" + listProductItemData[0][(int)EnumColorMeItem.ORDER_DATE], fnt);
            para.Alignment = Element.ALIGN_RIGHT;
            doc.Add(para);

            fnt = new Font(BaseFont.CreateFont(@"c:\windows\fonts\msgothic.ttc,1", BaseFont.IDENTITY_H, true), 8);
            para = new Paragraph("\n", fnt);
            doc.Add(para);

            fnt = new Font(BaseFont.CreateFont(@"c:\windows\fonts\msgothic.ttc,1", BaseFont.IDENTITY_H, true), 12);
            para = new Paragraph("この度は、当店をご利用いただき、誠にありがとうございます。", fnt);
            doc.Add(para);

            // 住所
            tbl = new Table(2);
            tbl.Width = 100;
            tbl.Widths = new float[] { 0.50f, 0.50f };
            tbl.DefaultHorizontalAlignment = Element.ALIGN_LEFT;
            tbl.Padding = 0;
            tbl.Spacing = 10;
            tbl.BorderColor = new iTextSharp.text.Color(0, 0, 0);

            //タイトル
            fnt = new Font(BaseFont.CreateFont(@"c:\windows\fonts\msgothic.ttc,1", BaseFont.IDENTITY_H, true), 14);

            cel = new Cell(new Phrase("ご注文者", fnt));
            tbl.AddCell(cel);
            cel = new Cell(new Phrase("お届け先", fnt));
            cel.VerticalAlignment = Element.ALIGN_TOP;
            tbl.AddCell(cel);
            // 要素
            fnt = new Font(BaseFont.CreateFont(@"c:\windows\fonts\msgothic.ttc,1", BaseFont.IDENTITY_H, true), 10);

            cel = new Cell(new Phrase("〒" + listProductItemData[0][(int)EnumColorMeItem.ORDER_POST_NO] + "\n" +
                                      listProductItemData[0][(int)EnumColorMeItem.ORDER_ADDRESS_1] +
                                      listProductItemData[0][(int)EnumColorMeItem.ORDER_ADDRESS_2] + "\n\n" +
                                      listProductItemData[0][(int)EnumColorMeItem.ORDER_NAME] + " 様\nTEL：" +
                                      listProductItemData[0][(int)EnumColorMeItem.ORDER_PHONE_NO], fnt));
            tbl.AddCell(cel);
            cel = new Cell(new Phrase("〒" + listProductItemData[0][(int)EnumColorMeItem.POST_NO] + "\n" +
                                      listProductItemData[0][(int)EnumColorMeItem.ADDRESS_1] +
                                      listProductItemData[0][(int)EnumColorMeItem.ADDRESS_2] + "\n\n" +
                                      listProductItemData[0][(int)EnumColorMeItem.NAME] + " 様\nTEL：" +
                                      listProductItemData[0][(int)EnumColorMeItem.PHONE_NO], fnt));
            tbl.AddCell(cel);

            // テーブル追加
            doc.Add(tbl);

            // 改行
            fnt = new Font(BaseFont.CreateFont(@"c:\windows\fonts\msgothic.ttc,1", BaseFont.IDENTITY_H, true), 8);
            para = new Paragraph("\n", fnt);
            doc.Add(para);

            // 総合計金額
            fnt = new Font(BaseFont.CreateFont(@"c:\windows\fonts\msgothic.ttc,1", BaseFont.IDENTITY_H, true), 18, iTextSharp.text.Font.UNDERLINE);
            para = new Paragraph("総合計金額(税込)　　" + int.Parse(listProductItemData[0][(int)EnumColorMeItem.TOTAL_PAY]).ToString("#,0") + " 円", fnt);
            doc.Add(para);

            // 商品情報
            tbl = new Table(5);
            tbl.Width = 100;
            tbl.Widths = new float[] { 0.10f, 0.55f, 0.10f, 0.10f, 0.15f };
            tbl.DefaultHorizontalAlignment = Element.ALIGN_LEFT;
            tbl.Padding = 0;
            tbl.Spacing = 5;
            tbl.BorderColor = new iTextSharp.text.Color(0, 0, 0);

            //タイトル
            fnt = new Font(BaseFont.CreateFont(@"c:\windows\fonts\msgothic.ttc,1", BaseFont.IDENTITY_H, true), 10);

            cel = new Cell(new Phrase("型番", fnt));
            tbl.AddCell(cel);
            cel = new Cell(new Phrase("商品名", fnt));
            tbl.AddCell(cel);
            cel = new Cell(new Phrase("単価", fnt));
            tbl.AddCell(cel);
            cel = new Cell(new Phrase("数量", fnt));
            tbl.AddCell(cel);
            cel = new Cell(new Phrase("金額", fnt));
            tbl.AddCell(cel);

            // 要素
            for (int i = 0; i < listProductItemData.Count; i++)
            {
                cel = new Cell(new Phrase(listProductItemData[i][(int)EnumColorMeItem.SKU], fnt));
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                tbl.AddCell(cel);

                cel = new Cell(new Phrase(listProductItemData[i][(int)EnumColorMeItem.PRODUCT_NAME], fnt));
                tbl.AddCell(cel);

                cel = new Cell(new Phrase(int.Parse(listProductItemData[i][(int)EnumColorMeItem.PRICE]).ToString("#,0") + " 円", fnt));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                tbl.AddCell(cel);

                cel = new Cell(new Phrase(listProductItemData[i][(int)EnumColorMeItem.UNIT], fnt));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                tbl.AddCell(cel);

                cel = new Cell(new Phrase(int.Parse(listProductItemData[i][(int)EnumColorMeItem.SUB_TOTAL]).ToString("#,0") + " 円", fnt));
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                tbl.AddCell(cel);
            }

            cel = new Cell(new Phrase("商品の合計金額(税込)", fnt));
            cel.Colspan = 4;
            cel.HorizontalAlignment = Element.ALIGN_RIGHT;
            tbl.AddCell(cel);

            cel = new Cell(new Phrase(int.Parse(listProductItemData[0][(int)EnumColorMeItem.TOTAL_PRODUCT_PRICE]).ToString("#,0") + " 円", fnt));
            cel.HorizontalAlignment = Element.ALIGN_RIGHT;
            tbl.AddCell(cel);

            cel = new Cell(new Phrase("送料合計(税込)", fnt));
            cel.Colspan = 4;
            cel.HorizontalAlignment = Element.ALIGN_RIGHT;
            tbl.AddCell(cel);

            cel = new Cell(new Phrase(int.Parse(listProductItemData[0][(int)EnumColorMeItem.SHIPPING_COST]).ToString("#,0") + " 円", fnt));
            cel.HorizontalAlignment = Element.ALIGN_RIGHT;
            tbl.AddCell(cel);

            cel = new Cell(new Phrase("決済手数料(税込)", fnt));
            cel.Colspan = 4;
            cel.HorizontalAlignment = Element.ALIGN_RIGHT;
            tbl.AddCell(cel);

            cel = new Cell(new Phrase(int.Parse(listProductItemData[0][(int)EnumColorMeItem.COMMISSION]).ToString("#,0") + " 円", fnt));
            cel.HorizontalAlignment = Element.ALIGN_RIGHT;
            tbl.AddCell(cel);

            // テーブル追加
            doc.Add(tbl);

            // 画像設定
            iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(new Uri(Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]) + "\\logo.jpg"));
            image.ScalePercent(50.0f);
            image.SetAbsolutePosition(350f, 3f);
            doc.Add(image);

            // 改ページ
            doc.NewPage();

            return true;
        }
        //--------------------------------------------------------------
        //--------------------------------------------------------------
    }
}
