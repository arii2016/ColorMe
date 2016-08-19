using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ColorMe
{
    //--------------------------------------------------------------
    #region 列挙型
    /// <summary>
    /// 読み込みCSVファイル種別
    /// </summary>
    public enum EnumReadCsvClass
    {
        /// <summary>カラーミー受注情報</summary>
        COLOR_ME = 0,

        MAX
    }
    /// <summary>
    /// カラーミー要素
    /// </summary>
    public enum EnumColorMeItem
    {
        /// <summary>売上ID</summary>
        ORDER_ID = 0,
        /// <summary>受注日</summary>
        ORDER_DATE = 1,
        /// <summary>購入者名前</summary>
        ORDER_NAME = 4,
        /// <summary>購入者郵便番号</summary>
        ORDER_POST_NO = 5,
        /// <summary>購入者住所1</summary>
        ORDER_ADDRESS_1 = 6,
        /// <summary>購入者住所2</summary>
        ORDER_ADDRESS_2 = 7,
        /// <summary>メールアドレス</summary>
        E_MAIL = 8,
        /// <summary>購入者電話番号</summary>
        ORDER_PHONE_NO = 9,
        /// <summary>商品の合計金額(税込)</summary>
        TOTAL_PRODUCT_PRICE = 10,
        /// <summary>消費税(商品合計に対する)</summary>
        TAX = 11,
        /// <summary>送料合計</summary>
        SHIPPING_COST = 12,
        /// <summary>決済手数料</summary>
        COMMISSION = 13,
        /// <summary>総合計金額</summary>
        TOTAL_PAY = 21,
        /// <summary>決済方法</summary>
        PAYMENT_METHOD = 22,
        /// <summary>名前</summary>
        NAME = 34,
        /// <summary>郵便番号</summary>
        POST_NO = 36,
        /// <summary>住所1</summary>
        ADDRESS_1 = 37,
        /// <summary>住所2</summary>
        ADDRESS_2 = 38,
        /// <summary>電話番号</summary>
        PHONE_NO = 39,
        /// <summary>配送方法</summary>
        DELIVERY_METHOD = 40,
        /// <summary>配送希望日</summary>
        DELIBERY_DATE = 41,
        /// <summary>配送希望時間</summary>
        DELIBERY_TIME = 42,
        /// <summary>型番</summary>
        SKU = 55,
        /// <summary>商品ID</summary>
        PRODUCT_ID = 56,
        /// <summary>商品名</summary>
        PRODUCT_NAME = 57,
        /// <summary>販売価格(消費税込)</summary>
        PRICE = 58,
        /// <summary>販売個数</summary>
        UNIT = 60,
        /// <summary>小計</summary>
        SUB_TOTAL = 61
    }
    /// <summary>
    /// ゆうプリR要素
    /// </summary>
    public enum EnumYpprItem
    {
        /// <summary>お客様側管理番号</summary>
        USER_ID = 0,
        /// <summary>発送予定日</summary>
        SHIPPING_SCHEDULE_DATE,
        /// <summary>発送予定時間</summary>
        SHIPPING_SCHEDULE_TIME,
        /// <summary>郵便種別</summary>
        POST_CLASS,
        /// <summary>支払元</summary>
        PAYMENT_SOURCE,
        /// <summary>送り状種別</summary>
        INVOICE_CLASS,
        /// <summary>お届け先郵便番号</summary>
        TRANSPORT_POST_NO,
        /// <summary>お届け先住所1</summary>
        TRANSPORT_ADDRESS_1,
        /// <summary>お届け先住所2</summary>
        TRANSPORT_ADDRESS_2,
        /// <summary>お届け先住所3</summary>
        TRANSPORT_ADDRESS_3,
        /// <summary>お届け先名称1</summary>
        TRANSPORT_NAME_1,
        /// <summary>お届け先名称2</summary>
        TRANSPORT_NAME_2,
        /// <summary>お届け先敬称</summary>
        TRANSPORT_TITLE,
        /// <summary>お届け先電話番号</summary>
        TRANSPORT_TEL,
        /// <summary>お届け先メール</summary>
        TRANSPORT_MAIL,
        /// <summary>発送元郵便番号</summary>
        ORIGIN_POST_NO,
        /// <summary>発送元住所1</summary>
        ORIGIN_ADDRESS_1,
        /// <summary>発送元住所2</summary>
        ORIGIN_ADDRESS_2,
        /// <summary>発送元住所3</summary>
        ORIGIN_ADDRESS_3,
        /// <summary>発送元名称1</summary>
        ORIGIN_NAME_1,
        /// <summary>発送元名称2</summary>
        ORIGIN_NAME_2,
        /// <summary>発送元敬称</summary>
        ORIGIN_TITLE,
        /// <summary>発送元電話番号</summary>
        ORIGIN_TEL,
        /// <summary>発送元メール</summary>
        ORIGIN_MAIL,
        /// <summary>こわれもの</summary>
        BREAKABLE_FLG,
        /// <summary>逆さま厳禁</summary>
        WAY_UP_FLG,
        /// <summary>下積み厳禁</summary>
        DO_NOT_STACK_FLG,
        /// <summary>厚さ</summary>
        THICKNESS,
        /// <summary>お届け日</summary>
        DELIBERY_DATE,
        /// <summary>お届け時間</summary>
        DELIBERY_TIME,
        /// <summary>フリー項目</summary>
        FREE_ITEM,
        /// <summary>代引金額</summary>
        COD_PAY,
        /// <summary>代引消費税</summary>
        COD_TAX,
        /// <summary>商品名</summary>
        PRODUCT_NAME,

        MAX
    }
    /// <summary>
    /// ヤマト要素
    /// </summary>
    public enum EnumYamatoItem
    {
        /// <summary>お客様側管理番号</summary>
        USER_ID = 0,
        /// <summary>送り状種別</summary>
        INVOICE_CLASS,
        /// <summary>NULL</summary>
        NULL_1,
        /// <summary>NULL</summary>
        NULL_2,
        /// <summary>発送予定時間</summary>
        SHIPPING_SCHEDULE_TIME,
        /// <summary>配達指定日</summary>
        DELIBERY_DATE,
        /// <summary>配達時間帯区分</summary>
        DELIBERY_TIME,
        /// <summary>NULL</summary>
        NULL_3,
        /// <summary>お届け先電話番号</summary>
        TRANSPORT_TEL,
        /// <summary>NULL</summary>
        NULL_4,
        /// <summary>お届け先郵便番号</summary>
        TRANSPORT_POST_NO,
        /// <summary>お届け先住所1</summary>
        TRANSPORT_ADDRESS_1,
        /// <summary>お届け先住所2</summary>
        TRANSPORT_ADDRESS_2,
        /// <summary>お届け先会社・部門名1</summary>
        TRANSPORT_COMPANY_1,
        /// <summary>お届け先会社・部門名2</summary>
        TRANSPORT_COMPANY_2,
        /// <summary>お届け先名</summary>
        TRANSPORT_NAME,
        /// <summary>NULL</summary>
        NULL_5,
        /// <summary>お届け先敬称</summary>
        TRANSPORT_TITLE,
        /// <summary>ご依頼主コード</summary>
        ORIGIN_CODE,
        /// <summary>ご依頼主電話番号</summary>
        ORIGIN_TEL,
        /// <summary>NULL</summary>
        NULL_6,
        /// <summary>ご依頼主郵便番号</summary>
        ORIGIN_POST_NO,
        /// <summary>ご依頼主住所1</summary>
        ORIGIN_ADDRESS_1,
        /// <summary>ご依頼主住所2</summary>
        ORIGIN_ADDRESS_2,
        /// <summary>ご依頼主名</summary>
        ORIGIN_NAME,
        /// <summary>NULL</summary>
        NULL_7,
        /// <summary>品名コード1</summary>
        PRODUCT_NAME_CODE_1,
        /// <summary>品名1</summary>
        PRODUCT_NAME_1,
        /// <summary>品名コード2</summary>
        PRODUCT_NAME_CODE_2,
        /// <summary>品名2</summary>
        PRODUCT_NAME_2,
        /// <summary>荷扱い1</summary>
        FREIGHT_HANDLING_1,
        /// <summary>荷扱い2</summary>
        FREIGHT_HANDLING_2,
        /// <summary>記事</summary>
        ARTICLE,
        /// <summary>代引金額</summary>
        COD_PAY,
        /// <summary>代引消費税</summary>
        COD_TAX,
        /// <summary>NULL</summary>
        NULL_8,
        /// <summary>NULL</summary>
        NULL_9,
        /// <summary>発行枚数</summary>
        POST_NUM,
        /// <summary>個数口枠の印字</summary>
        NUMBER_FRAME,
        /// <summary>ご請求先顧客コード</summary>
        BILLING_CODE,
        /// <summary>NULL</summary>
        NULL_10,
        /// <summary>運賃管理番号</summary>
        FARE_NO,

        MAX
    }
    #endregion
    //--------------------------------------------------------------
    /// <summary>
    /// 定数宣言クラス
    /// </summary>
    class CDefine
    {
        //--------------------------------------------------------------
        #region 文字列定数
        /// <summary>読み込みCSVファイル種別文字列定数</summary>
        public static string[] READ_CSV_CLASS_CONF_STR = new string[(int)EnumReadCsvClass.MAX];
        #endregion
        //--------------------------------------------------------------
        /// <summary>コンストラクタ</summary>
        static CDefine()
        {
            //--------------------------------------------------------------
            #region 読み込みCSVファイル判断文字列定数
            READ_CSV_CLASS_CONF_STR[(int)EnumReadCsvClass.COLOR_ME] = "\"売上ID\"";
            #endregion
            //--------------------------------------------------------------
        }
        //--------------------------------------------------------------
    }
    //--------------------------------------------------------------
}
