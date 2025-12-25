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
