using FluentAssertions;
using VbaMcpServer.Helpers;

namespace VbaMcpServer.Tests.Helpers;

public class ComHelperTests
{
    [Theory]
    [InlineData(unchecked((int)0x800401E3), "Application is not running")]
    [InlineData(unchecked((int)0x80070005), "Access denied to VBA project")]
    [InlineData(unchecked((int)0x80020003), "Module or member not found")]
    [InlineData(unchecked((int)0x12345678), "COM error: 0x12345678")]
    public void GetErrorMessage_ReturnsCorrectMessage(int hresult, string expected)
    {
        // Act
        var result = ComErrorCodes.GetErrorMessage(hresult);

        // Assert
        result.Should().Be(expected);
    }

    [Fact]
    public void IsApplicationUnavailable_ReturnsTrueForMK_E_UNAVAILABLE()
    {
        // Arrange
        int hresult = unchecked((int)0x800401E3);

        // Act
        var result = ComErrorCodes.IsApplicationUnavailable(hresult);

        // Assert
        result.Should().BeTrue();
    }

    [Fact]
    public void IsApplicationUnavailable_ReturnsFalseForOtherErrors()
    {
        // Arrange
        int hresult = unchecked((int)0x80070005);

        // Act
        var result = ComErrorCodes.IsApplicationUnavailable(hresult);

        // Assert
        result.Should().BeFalse();
    }

    [Theory]
    [InlineData(unchecked((int)0x80070005), true)]
    [InlineData(unchecked((int)0x800401E3), false)]
    [InlineData(unchecked((int)0x80020003), false)]
    public void IsVbaAccessError_ReturnsCorrectResult(int hresult, bool expected)
    {
        // Act
        var result = ComErrorCodes.IsVbaAccessError(hresult);

        // Assert
        result.Should().Be(expected);
    }

    [Theory]
    [InlineData(unchecked((int)0x80020003), true)]
    [InlineData(unchecked((int)0x8002000B), true)]
    [InlineData(unchecked((int)0x80070005), false)]
    public void IsNotFoundError_ReturnsCorrectResult(int hresult, bool expected)
    {
        // Act
        var result = ComErrorCodes.IsNotFoundError(hresult);

        // Assert
        result.Should().Be(expected);
    }
}
