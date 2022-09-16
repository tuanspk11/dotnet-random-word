//*****************************************************************************
//* ALL RIGHTS RESERVED. COPYRIGHT (C) ソリマチ株式会社                       *
//*****************************************************************************
//* File Name    : ExportFile.cs                                              *
//* Function     : CSVファイル出力処理                                        *
//* System Name  : 会計事務所クラウド                                         *
//* Create       : 2020/11/18 TrinhNguyen                                     *
//* Update       :                                                            *
//* Comment      : 　                                                         *
//*****************************************************************************

using ExportAdvisors.Mgrs;
using ExportAdvisors.Models.DB;
using ExportAdvisors.Models.Storage.Tables;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExportAdvisors.Funcs
{
    /// <summary> CSVファイル出力処理 </summary>
    internal class ExportFile
    {
        /// <summary> ファイル出力パス </summary>
        private static string basePath { get; set; }

        /// <summary> CSV出力処理 </summary>
        /// <param name="log">処理ログ出力</param>
        public static void OutputCSV(LogMgr log)
        {
            log.Output("CSV出力開始", LogMgr.Process.Start, "");

            PortalStorageAccess storageMgr = new PortalStorageAccess();
            basePath = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\ExportCSV";

            if (!System.IO.Directory.Exists(basePath))
            {
                System.IO.Directory.CreateDirectory(basePath);
            }

            // ストレージからファイル情報を取得する
            List<SharefilesData> lstSharefiles = storageMgr.GetListSharefiles();
            if (lstSharefiles.Count() == 0)
            {
                // ログ出力(0件)
                log.Output("sharefilesは0件のため出力なし", LogMgr.Process.Start, "");
            }
            else
            {               
                DBRepository dbMgr = new DBRepository();
                // CSV出力項目作成
                dbMgr.MakeResult(ref lstSharefiles);

                // CSVファイル作成
                MakeCSVFile(lstSharefiles);
            }
            log.Output("CSV出力終了", LogMgr.Process.End, "");
        }

        /// <summary>
        /// CSVファイルを作成する
        /// </summary>
        /// <param name="lstSharefiles"> CSV情報</param>
        private static void MakeCSVFile(List<SharefilesData> lstSharefiles)
        {
            //CSVファイルに書き込むときに使うEncoding
            System.Text.Encoding enc = System.Text.Encoding.GetEncoding("UTF-8");

            //CSVの出力先パス                
            //ファイル名
            string FileName = "\\顧問先一覧_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";
            string csvPath = basePath + FileName;
            if (System.IO.File.Exists(csvPath))
            {   // 存在時はもともとのファイルを削除
                System.IO.File.Delete(csvPath);
            }
            //書き込むファイルを開く
            System.IO.StreamWriter sr = new System.IO.StreamWriter(csvPath, false, enc);

            //ヘッダー書き込み
            string headerStr = WriteContentRow("事務所コード",
                                                "事務所名",
                                                "顧問先名",
                                                "サービスシリアルナンバー",
                                                "作成日時",
                                                "最終更新日時");
            sr.Write(headerStr);
            sr.Write("\r\n");

            foreach (var REF in lstSharefiles)
            {
                string contentRow = WriteContentRow(REF.KikCd,
                                                    REF.KikNm,
                                                    REF.KnyNm,
                                                    REF.ACC_serialkey,
                                                    REF.CreateDate.ToString(),
                                                    REF.UpdateDate.ToString());
                //書き込み
                sr.Write(contentRow);
                //改行する
                sr.Write("\r\n");
            }
            //閉じる
            sr.Close();
        }

        /// <summary>
        /// １行書き込み
        /// </summary>
        /// <param name="KikCd">事務所コード</param>
        /// <param name="KikNm">事務所名</param>
        /// <param name="KnyNm">顧問先名</param>
        /// <param name="ACC_serialkey">サービスシリアルナンバー</param>
        /// <param name="CreateDate">作成日時</param>
        /// <param name="UpdateDate">最終更新日時</param>
        /// <returns></returns>
        private static string WriteContentRow(string KikCd, string KikNm, string KnyNm, string ACC_serialkey, string CreateDate, string UpdateDate)
        {
            return String.Format("{0},{1},{2},{3},{4},{5}",
                                EncloseDoubleQuotes(KikCd),
                                EncloseDoubleQuotes(KikNm),
                                EncloseDoubleQuotes(KnyNm),
                                EncloseDoubleQuotes(ACC_serialkey),
                                EncloseDoubleQuotes(CreateDate),
                                EncloseDoubleQuotes(UpdateDate));
        }

        /// <summary>
        /// 文字列をダブルクォートで囲む
        /// </summary>
        private static string EncloseDoubleQuotes(string field)
        {
            if (field.IndexOf('"') > -1)
            {
                //"を""とする
                field = field.Replace("\"", "\"\"");
            }
            return "\"" + field + "\"";
        }
    }
}
