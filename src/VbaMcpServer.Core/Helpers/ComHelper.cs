using System.Runtime.InteropServices;

namespace VbaMcpServer.Helpers;

/// <summary>
/// Helper class for COM interop operations in .NET 8+
/// Provides P/Invoke implementations for deprecated Marshal methods
/// </summary>
public static class ComHelper
{
    /// <summary>
    /// Gets a running instance of a COM object by ProgID
    /// Replacement for Marshal.GetActiveObject which is not supported in .NET 8+
    /// </summary>
    /// <param name="progId">The programmatic identifier (ProgID) of the COM object</param>
    /// <returns>The running COM object instance</returns>
    /// <exception cref="COMException">Thrown when the object is not found or cannot be retrieved</exception>
    public static object GetActiveObject(string progId)
    {
        // Get CLSID from ProgID
        int hr = CLSIDFromProgID(progId, out Guid clsid);
        if (hr < 0)
        {
            Marshal.ThrowExceptionForHR(hr);
        }

        // Get the active object from Running Object Table (ROT)
        hr = GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
        if (hr < 0)
        {
            Marshal.ThrowExceptionForHR(hr);
        }

        return obj;
    }

    /// <summary>
    /// Retrieves the CLSID associated with the specified ProgID
    /// </summary>
    [DllImport("ole32.dll", PreserveSig = true)]
    private static extern int CLSIDFromProgID(
        [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
        out Guid pclsid);

    /// <summary>
    /// Retrieves a pointer to a running object registered in the Running Object Table (ROT)
    /// </summary>
    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(
        ref Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
}

/// <summary>
/// Common COM error codes
/// </summary>
public static class ComErrorCodes
{
    /// <summary>
    /// Object is not available (e.g., Excel is not running)
    /// </summary>
    public const int MK_E_UNAVAILABLE = unchecked((int)0x800401E3);

    /// <summary>
    /// Access denied
    /// </summary>
    public const int E_ACCESSDENIED = unchecked((int)0x80070005);

    /// <summary>
    /// Unspecified failure
    /// </summary>
    public const int E_FAIL = unchecked((int)0x80004005);

    /// <summary>
    /// Member not found (e.g., module not found)
    /// </summary>
    public const int DISP_E_MEMBERNOTFOUND = unchecked((int)0x80020003);

    /// <summary>
    /// Invalid index (e.g., subscript out of range)
    /// </summary>
    public const int DISP_E_BADINDEX = unchecked((int)0x8002000B);

    /// <summary>
    /// Type mismatch
    /// </summary>
    public const int DISP_E_TYPEMISMATCH = unchecked((int)0x80020005);

    /// <summary>
    /// Gets a user-friendly error message for a COM HRESULT
    /// </summary>
    /// <param name="hresult">The HRESULT error code</param>
    /// <returns>A user-friendly error message</returns>
    public static string GetErrorMessage(int hresult)
    {
        return hresult switch
        {
            MK_E_UNAVAILABLE => "Application is not running",
            E_ACCESSDENIED => "Access denied to VBA project",
            DISP_E_MEMBERNOTFOUND => "Module or member not found",
            DISP_E_BADINDEX => "Invalid index or subscript out of range",
            DISP_E_TYPEMISMATCH => "Type mismatch in COM operation",
            E_FAIL => "COM operation failed",
            _ => $"COM error: 0x{hresult:X8}"
        };
    }

    /// <summary>
    /// Checks if the HRESULT indicates a VBA project access issue
    /// </summary>
    public static bool IsVbaAccessError(int hresult)
    {
        return hresult == E_ACCESSDENIED;
    }

    /// <summary>
    /// Checks if the HRESULT indicates the application is not running
    /// </summary>
    public static bool IsApplicationUnavailable(int hresult)
    {
        return hresult == MK_E_UNAVAILABLE;
    }

    /// <summary>
    /// Checks if the HRESULT indicates a module/member not found error
    /// </summary>
    public static bool IsNotFoundError(int hresult)
    {
        return hresult == DISP_E_MEMBERNOTFOUND || hresult == DISP_E_BADINDEX;
    }
}
