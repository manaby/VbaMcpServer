using FluentAssertions;
using VbaMcpServer.Exceptions;

namespace VbaMcpServer.Tests.Exceptions;

/// <summary>
/// Tests for Access-related exception classes
/// </summary>
public class AccessExceptionsTests
{
    [Fact]
    public void FormNotFoundException_ShouldContainFormName()
    {
        // Arrange
        var formName = "MainForm";
        var filePath = "C:\\test.accdb";

        // Act
        var exception = new FormNotFoundException(formName, filePath);

        // Assert
        exception.Message.Should().Contain(formName);
        exception.ErrorCode.Should().Be("FORM_NOT_FOUND");
        exception.FilePath.Should().Be(filePath);
        exception.Should().BeAssignableTo<VbaMcpException>();
    }

    [Fact]
    public void FormNotFoundException_WithInnerException_ShouldPreserveInnerException()
    {
        // Arrange
        var formName = "MainForm";
        var innerException = new InvalidOperationException("Test inner exception");

        // Act
        var exception = new FormNotFoundException(formName, innerException);

        // Assert
        exception.InnerException.Should().Be(innerException);
        exception.Message.Should().Contain(formName);
    }

    [Fact]
    public void ReportNotFoundException_ShouldContainReportName()
    {
        // Arrange
        var reportName = "SalesReport";
        var filePath = "C:\\test.accdb";

        // Act
        var exception = new ReportNotFoundException(reportName, filePath);

        // Assert
        exception.Message.Should().Contain(reportName);
        exception.ErrorCode.Should().Be("REPORT_NOT_FOUND");
        exception.FilePath.Should().Be(filePath);
        exception.Should().BeAssignableTo<VbaMcpException>();
    }

    [Fact]
    public void ReportNotFoundException_WithInnerException_ShouldPreserveInnerException()
    {
        // Arrange
        var reportName = "SalesReport";
        var innerException = new InvalidOperationException("Test inner exception");

        // Act
        var exception = new ReportNotFoundException(reportName, innerException);

        // Assert
        exception.InnerException.Should().Be(innerException);
        exception.Message.Should().Contain(reportName);
    }

    [Fact]
    public void ControlNotFoundException_ShouldContainControlAndObjectName()
    {
        // Arrange
        var controlName = "btnSubmit";
        var objectName = "Form_MainForm";
        var filePath = "C:\\test.accdb";

        // Act
        var exception = new ControlNotFoundException(controlName, objectName, filePath);

        // Assert
        exception.Message.Should().Contain(controlName);
        exception.Message.Should().Contain(objectName);
        exception.ObjectName.Should().Be(objectName);
        exception.ErrorCode.Should().Be("CONTROL_NOT_FOUND");
        exception.FilePath.Should().Be(filePath);
        exception.Should().BeAssignableTo<VbaMcpException>();
    }

    [Fact]
    public void ControlNotFoundException_WithInnerException_ShouldPreserveInnerException()
    {
        // Arrange
        var controlName = "btnSubmit";
        var objectName = "Form_MainForm";
        var innerException = new InvalidOperationException("Test inner exception");

        // Act
        var exception = new ControlNotFoundException(controlName, objectName, innerException);

        // Assert
        exception.InnerException.Should().Be(innerException);
        exception.ObjectName.Should().Be(objectName);
    }

    [Fact]
    public void PropertyNotFoundException_ShouldContainPropertyAndControlName()
    {
        // Arrange
        var propertyName = "Caption";
        var controlName = "btnSubmit";
        var filePath = "C:\\test.accdb";

        // Act
        var exception = new PropertyNotFoundException(propertyName, controlName, filePath);

        // Assert
        exception.Message.Should().Contain(propertyName);
        exception.Message.Should().Contain(controlName);
        exception.ControlName.Should().Be(controlName);
        exception.ErrorCode.Should().Be("PROPERTY_NOT_FOUND");
        exception.FilePath.Should().Be(filePath);
        exception.Should().BeAssignableTo<VbaMcpException>();
    }

    [Fact]
    public void PropertyNotFoundException_WithInnerException_ShouldPreserveInnerException()
    {
        // Arrange
        var propertyName = "Caption";
        var controlName = "btnSubmit";
        var innerException = new InvalidOperationException("Test inner exception");

        // Act
        var exception = new PropertyNotFoundException(propertyName, controlName, innerException);

        // Assert
        exception.InnerException.Should().Be(innerException);
        exception.ControlName.Should().Be(controlName);
    }

    [Fact]
    public void PropertyReadOnlyException_ShouldContainPropertyAndControlName()
    {
        // Arrange
        var propertyName = "Name";
        var controlName = "btnSubmit";
        var filePath = "C:\\test.accdb";

        // Act
        var exception = new PropertyReadOnlyException(propertyName, controlName, filePath);

        // Assert
        exception.Message.Should().Contain(propertyName);
        exception.Message.Should().Contain(controlName);
        exception.Message.Should().Contain("read-only");
        exception.ControlName.Should().Be(controlName);
        exception.ErrorCode.Should().Be("PROPERTY_READ_ONLY");
        exception.FilePath.Should().Be(filePath);
        exception.Should().BeAssignableTo<VbaMcpException>();
    }

    [Fact]
    public void PropertyReadOnlyException_WithInnerException_ShouldPreserveInnerException()
    {
        // Arrange
        var propertyName = "Name";
        var controlName = "btnSubmit";
        var innerException = new InvalidOperationException("Test inner exception");

        // Act
        var exception = new PropertyReadOnlyException(propertyName, controlName, innerException);

        // Assert
        exception.InnerException.Should().Be(innerException);
        exception.ControlName.Should().Be(controlName);
    }

    [Fact]
    public void InvalidPropertyValueException_ShouldContainPropertyAndValue()
    {
        // Arrange
        var propertyName = "Width";
        var value = "invalid";
        var filePath = "C:\\test.accdb";

        // Act
        var exception = new InvalidPropertyValueException(propertyName, value, filePath);

        // Assert
        exception.Message.Should().Contain(propertyName);
        exception.Message.Should().Contain(value);
        exception.PropertyName.Should().Be(propertyName);
        exception.ProvidedValue.Should().Be(value);
        exception.ErrorCode.Should().Be("INVALID_VALUE");
        exception.FilePath.Should().Be(filePath);
        exception.Should().BeAssignableTo<VbaMcpException>();
    }

    [Fact]
    public void InvalidPropertyValueException_WithInnerException_ShouldPreserveInnerException()
    {
        // Arrange
        var propertyName = "Width";
        var value = "invalid";
        var innerException = new InvalidOperationException("Test inner exception");

        // Act
        var exception = new InvalidPropertyValueException(propertyName, value, innerException);

        // Assert
        exception.InnerException.Should().Be(innerException);
        exception.PropertyName.Should().Be(propertyName);
        exception.ProvidedValue.Should().Be(value);
    }

    [Fact]
    public void AllAccessExceptions_ShouldInheritFromVbaMcpException()
    {
        // Arrange & Act
        var formException = new FormNotFoundException("test");
        var reportException = new ReportNotFoundException("test");
        var controlException = new ControlNotFoundException("control", "object");
        var propertyNotFoundException = new PropertyNotFoundException("prop", "control");
        var propertyReadOnlyException = new PropertyReadOnlyException("prop", "control");
        var invalidPropertyException = new InvalidPropertyValueException("prop", "value");

        // Assert
        formException.Should().BeAssignableTo<VbaMcpException>();
        reportException.Should().BeAssignableTo<VbaMcpException>();
        controlException.Should().BeAssignableTo<VbaMcpException>();
        propertyNotFoundException.Should().BeAssignableTo<VbaMcpException>();
        propertyReadOnlyException.Should().BeAssignableTo<VbaMcpException>();
        invalidPropertyException.Should().BeAssignableTo<VbaMcpException>();
    }

    [Fact]
    public void AllAccessExceptions_ShouldHaveUniqueErrorCodes()
    {
        // Arrange & Act
        var formException = new FormNotFoundException("test");
        var reportException = new ReportNotFoundException("test");
        var controlException = new ControlNotFoundException("control", "object");
        var propertyNotFoundException = new PropertyNotFoundException("prop", "control");
        var propertyReadOnlyException = new PropertyReadOnlyException("prop", "control");
        var invalidPropertyException = new InvalidPropertyValueException("prop", "value");

        var errorCodes = new[]
        {
            formException.ErrorCode,
            reportException.ErrorCode,
            controlException.ErrorCode,
            propertyNotFoundException.ErrorCode,
            propertyReadOnlyException.ErrorCode,
            invalidPropertyException.ErrorCode
        };

        // Assert
        errorCodes.Should().OnlyHaveUniqueItems();
    }
}
