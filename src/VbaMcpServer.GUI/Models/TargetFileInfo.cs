namespace VbaMcpServer.GUI.Models;

/// <summary>
/// ファイル種別 (Excel or Access)
/// </summary>
public enum FileType
{
    Unknown,
    Excel,     // .xlsm, .xlsx, .xlsb, .xls
    Access     // .accdb, .mdb
}

/// <summary>
/// 対象ファイルの情報
/// </summary>
public class TargetFileInfo
{
    /// <summary>
    /// ファイルの絶対パス
    /// </summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>
    /// ファイル種別 (Excel or Access)
    /// </summary>
    public FileType FileType { get; set; }

    /// <summary>
    /// ファイルが開いているかどうか
    /// </summary>
    public bool IsOpen { get; set; }

    /// <summary>
    /// 開いているアプリケーションのプロセスID
    /// </summary>
    public int? ProcessId { get; set; }

    /// <summary>
    /// 最終確認日時
    /// </summary>
    public DateTime LastChecked { get; set; }
}
