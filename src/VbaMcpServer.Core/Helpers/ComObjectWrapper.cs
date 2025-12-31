using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;

namespace VbaMcpServer.Helpers;

/// <summary>
/// COM参照を自動的に解放するラッパー
/// IDisposableパターンでusingステートメントと組み合わせて使用
/// </summary>
/// <typeparam name="T">COMオブジェクトの型</typeparam>
public sealed class ComObjectWrapper<T> : IDisposable where T : class
{
    private T? _comObject;
    private readonly ILogger? _logger;
    private bool _disposed = false;

    public ComObjectWrapper(T? comObject, ILogger? logger = null)
    {
        _comObject = comObject;
        _logger = logger;
    }

    /// <summary>
    /// ラップされたCOMオブジェクトを取得
    /// </summary>
    public T? Value => _comObject;

    public void Dispose()
    {
        if (_disposed)
            return;

        if (_comObject != null)
        {
            try
            {
                // COM参照カウントをデクリメント
                int refCount = Marshal.ReleaseComObject(_comObject);
                _logger?.LogTrace("Released COM object: {Type}, RefCount: {RefCount}",
                    typeof(T).Name, refCount);
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "Failed to release COM object: {Type}", typeof(T).Name);
            }
            finally
            {
                _comObject = null;
            }
        }

        _disposed = true;
    }
}

/// <summary>
/// COM列挙を安全に処理するヘルパークラス
/// </summary>
public static class ComEnumerableHelper
{
    /// <summary>
    /// COM列挙の各アイテムを自動解放しながら処理
    /// </summary>
    /// <typeparam name="T">COMオブジェクトの型</typeparam>
    /// <param name="collection">COM列挙コレクション</param>
    /// <param name="action">各アイテムに対するアクション</param>
    /// <param name="logger">ロガー（オプション）</param>
    public static void ForEach<T>(
        object collection,
        Action<T> action,
        ILogger? logger = null) where T : class
    {
        // コレクション自体も解放
        using var collectionWrapper = new ComObjectWrapper<object>(collection, logger);

        foreach (var item in (System.Collections.IEnumerable)collection)
        {
            // 各アイテムを自動解放
            using var itemWrapper = new ComObjectWrapper<T>(item as T, logger);
            if (itemWrapper.Value != null)
            {
                action(itemWrapper.Value);
            }
        }
    }

    /// <summary>
    /// COM列挙をリストに変換（各アイテムを自動解放）
    /// </summary>
    /// <typeparam name="TItem">元のCOMオブジェクト型</typeparam>
    /// <typeparam name="TResult">変換後の型</typeparam>
    /// <param name="collection">COM列挙コレクション</param>
    /// <param name="selector">変換関数</param>
    /// <param name="logger">ロガー（オプション）</param>
    /// <returns>変換結果のリスト</returns>
    public static List<TResult> ToList<TItem, TResult>(
        object collection,
        Func<TItem, TResult> selector,
        ILogger? logger = null) where TItem : class
    {
        var results = new List<TResult>();

        ForEach<TItem>(collection, item =>
        {
            results.Add(selector(item));
        }, logger);

        return results;
    }
}
