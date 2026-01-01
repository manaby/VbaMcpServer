using FluentAssertions;
using VbaMcpServer.Models;
using System.Text.Json;

namespace VbaMcpServer.Tests.Models;

/// <summary>
/// Tests for Access-related model classes
/// </summary>
public class AccessModelsTests
{
    #region FormControlInfo Tests

    [Fact]
    public void FormControlInfo_ShouldInitializeWithRequiredProperties()
    {
        // Arrange & Act
        var controlInfo = new FormControlInfo
        {
            Name = "btnSubmit",
            ControlType = "CommandButton",
            ControlTypeId = 104,
            Section = "Detail",
            SectionId = 0,
            Left = 100,
            Top = 200,
            Width = 1000,
            Height = 500,
            Visible = true
        };

        // Assert
        controlInfo.Name.Should().Be("btnSubmit");
        controlInfo.ControlType.Should().Be("CommandButton");
        controlInfo.ControlTypeId.Should().Be(104);
        controlInfo.Section.Should().Be("Detail");
        controlInfo.SectionId.Should().Be(0);
        controlInfo.Left.Should().Be(100);
        controlInfo.Top.Should().Be(200);
        controlInfo.Width.Should().Be(1000);
        controlInfo.Height.Should().Be(500);
        controlInfo.Visible.Should().BeTrue();
    }

    [Fact]
    public void FormControlInfo_ShouldHandleOptionalProperties()
    {
        // Arrange & Act
        var controlInfo = new FormControlInfo
        {
            Name = "txtName",
            ControlType = "TextBox",
            ControlTypeId = 109,
            Section = "Detail",
            SectionId = 0,
            Left = 100,
            Top = 200,
            Width = 1000,
            Height = 500,
            Visible = true,
            Enabled = false,
            TabIndex = 5,
            ControlSource = "CustomerName",
            Parent = "Form_MainForm",
            SourceObject = null
        };

        // Assert
        controlInfo.Enabled.Should().BeFalse();
        controlInfo.TabIndex.Should().Be(5);
        controlInfo.ControlSource.Should().Be("CustomerName");
        controlInfo.Parent.Should().Be("Form_MainForm");
        controlInfo.SourceObject.Should().BeNull();
    }

    [Fact]
    public void FormControlInfo_ShouldSerializeToJson()
    {
        // Arrange
        var controlInfo = new FormControlInfo
        {
            Name = "btnSubmit",
            ControlType = "CommandButton",
            ControlTypeId = 104,
            Section = "Detail",
            SectionId = 0,
            Left = 100,
            Top = 200,
            Width = 1000,
            Height = 500,
            Visible = true
        };

        // Act
        var json = JsonSerializer.Serialize(controlInfo);
        var deserialized = JsonSerializer.Deserialize<FormControlInfo>(json);

        // Assert
        deserialized.Should().NotBeNull();
        deserialized!.Name.Should().Be(controlInfo.Name);
        deserialized.ControlType.Should().Be(controlInfo.ControlType);
        deserialized.ControlTypeId.Should().Be(controlInfo.ControlTypeId);
        deserialized.Visible.Should().Be(controlInfo.Visible);
    }

    #endregion

    #region ReportControlInfo Tests

    [Fact]
    public void ReportControlInfo_ShouldInitializeWithRequiredProperties()
    {
        // Arrange & Act
        var controlInfo = new ReportControlInfo
        {
            Name = "lblTitle",
            ControlType = "Label",
            ControlTypeId = 100,
            Section = "PageHeader",
            SectionId = 2,
            Left = 100,
            Top = 200,
            Width = 2000,
            Height = 300,
            Visible = true
        };

        // Assert
        controlInfo.Name.Should().Be("lblTitle");
        controlInfo.ControlType.Should().Be("Label");
        controlInfo.ControlTypeId.Should().Be(100);
        controlInfo.Section.Should().Be("PageHeader");
        controlInfo.SectionId.Should().Be(2);
        controlInfo.Left.Should().Be(100);
        controlInfo.Top.Should().Be(200);
        controlInfo.Width.Should().Be(2000);
        controlInfo.Height.Should().Be(300);
        controlInfo.Visible.Should().BeTrue();
    }

    [Fact]
    public void ReportControlInfo_ShouldHandleOptionalProperties()
    {
        // Arrange & Act
        var controlInfo = new ReportControlInfo
        {
            Name = "txtTotal",
            ControlType = "TextBox",
            ControlTypeId = 109,
            Section = "Detail",
            SectionId = 0,
            Left = 100,
            Top = 200,
            Width = 1000,
            Height = 500,
            Visible = true,
            ControlSource = "=Sum([Amount])",
            Parent = "Report_Sales"
        };

        // Assert
        controlInfo.ControlSource.Should().Be("=Sum([Amount])");
        controlInfo.Parent.Should().Be("Report_Sales");
    }

    [Fact]
    public void ReportControlInfo_ShouldSerializeToJson()
    {
        // Arrange
        var controlInfo = new ReportControlInfo
        {
            Name = "lblTitle",
            ControlType = "Label",
            ControlTypeId = 100,
            Section = "PageHeader",
            SectionId = 2,
            Left = 100,
            Top = 200,
            Width = 2000,
            Height = 300,
            Visible = true
        };

        // Act
        var json = JsonSerializer.Serialize(controlInfo);
        var deserialized = JsonSerializer.Deserialize<ReportControlInfo>(json);

        // Assert
        deserialized.Should().NotBeNull();
        deserialized!.Name.Should().Be(controlInfo.Name);
        deserialized.ControlType.Should().Be(controlInfo.ControlType);
        deserialized.Section.Should().Be(controlInfo.Section);
    }

    #endregion

    #region ControlPropertyInfo Tests

    [Fact]
    public void ControlPropertyInfo_ShouldInitializeWithRequiredProperties()
    {
        // Arrange
        var properties = new Dictionary<string, object?>
        {
            { "Caption", "Submit" },
            { "Width", 1000 },
            { "Height", 500 },
            { "Enabled", true }
        };

        // Act
        var propertyInfo = new ControlPropertyInfo
        {
            File = "C:\\test.accdb",
            ObjectName = "Form_MainForm",
            ControlName = "btnSubmit",
            ControlType = "CommandButton",
            Properties = properties
        };

        // Assert
        propertyInfo.File.Should().Be("C:\\test.accdb");
        propertyInfo.ObjectName.Should().Be("Form_MainForm");
        propertyInfo.ControlName.Should().Be("btnSubmit");
        propertyInfo.ControlType.Should().Be("CommandButton");
        propertyInfo.Properties.Should().HaveCount(4);
        propertyInfo.Properties["Caption"].Should().Be("Submit");
        propertyInfo.Properties["Width"].Should().Be(1000);
    }

    [Fact]
    public void ControlPropertyInfo_ShouldHandleNullPropertyValues()
    {
        // Arrange
        var properties = new Dictionary<string, object?>
        {
            { "ControlSource", null },
            { "Caption", "Test" }
        };

        // Act
        var propertyInfo = new ControlPropertyInfo
        {
            File = "C:\\test.accdb",
            ObjectName = "Form_MainForm",
            ControlName = "txtName",
            ControlType = "TextBox",
            Properties = properties
        };

        // Assert
        propertyInfo.Properties["ControlSource"].Should().BeNull();
        propertyInfo.Properties["Caption"].Should().Be("Test");
    }

    [Fact]
    public void ControlPropertyInfo_ShouldSerializeToJson()
    {
        // Arrange
        var properties = new Dictionary<string, object?>
        {
            { "Caption", "Submit" },
            { "Enabled", true }
        };

        var propertyInfo = new ControlPropertyInfo
        {
            File = "C:\\test.accdb",
            ObjectName = "Form_MainForm",
            ControlName = "btnSubmit",
            ControlType = "CommandButton",
            Properties = properties
        };

        // Act
        var json = JsonSerializer.Serialize(propertyInfo);
        var deserialized = JsonSerializer.Deserialize<ControlPropertyInfo>(json);

        // Assert
        deserialized.Should().NotBeNull();
        deserialized!.File.Should().Be(propertyInfo.File);
        deserialized.ObjectName.Should().Be(propertyInfo.ObjectName);
        deserialized.ControlName.Should().Be(propertyInfo.ControlName);
        deserialized.Properties.Should().HaveCount(2);
    }

    #endregion

    #region SetPropertyResult Tests

    [Fact]
    public void SetPropertyResult_ShouldIndicateSuccess()
    {
        // Arrange & Act
        var result = new SetPropertyResult
        {
            Success = true,
            File = "C:\\test.accdb",
            ObjectName = "Form_MainForm",
            ControlName = "btnSubmit",
            PropertyName = "Caption",
            PreviousValue = "Old Caption",
            NewValue = "New Caption"
        };

        // Assert
        result.Success.Should().BeTrue();
        result.File.Should().Be("C:\\test.accdb");
        result.ObjectName.Should().Be("Form_MainForm");
        result.ControlName.Should().Be("btnSubmit");
        result.PropertyName.Should().Be("Caption");
        result.PreviousValue.Should().Be("Old Caption");
        result.NewValue.Should().Be("New Caption");
        result.Error.Should().BeNull();
        result.ErrorCode.Should().BeNull();
    }

    [Fact]
    public void SetPropertyResult_ShouldIndicateFailure()
    {
        // Arrange & Act
        var result = new SetPropertyResult
        {
            Success = false,
            File = "C:\\test.accdb",
            ObjectName = "Form_MainForm",
            ControlName = "btnSubmit",
            PropertyName = "InvalidProperty",
            Error = "Property not found",
            ErrorCode = "PROPERTY_NOT_FOUND"
        };

        // Assert
        result.Success.Should().BeFalse();
        result.Error.Should().Be("Property not found");
        result.ErrorCode.Should().Be("PROPERTY_NOT_FOUND");
        result.PreviousValue.Should().BeNull();
        result.NewValue.Should().BeNull();
    }

    [Fact]
    public void SetPropertyResult_ShouldSerializeToJson()
    {
        // Arrange
        var result = new SetPropertyResult
        {
            Success = true,
            File = "C:\\test.accdb",
            ObjectName = "Form_MainForm",
            ControlName = "btnSubmit",
            PropertyName = "Caption",
            PreviousValue = "Old",
            NewValue = "New"
        };

        // Act
        var json = JsonSerializer.Serialize(result);
        var deserialized = JsonSerializer.Deserialize<SetPropertyResult>(json);

        // Assert
        deserialized.Should().NotBeNull();
        deserialized!.Success.Should().BeTrue();
        deserialized.File.Should().Be(result.File);
        deserialized.PropertyName.Should().Be(result.PropertyName);
    }

    [Fact]
    public void SetPropertyResult_ShouldHandleComplexPropertyValues()
    {
        // Arrange
        var complexValue = new { Width = 1000, Height = 500 };

        // Act
        var result = new SetPropertyResult
        {
            Success = true,
            PropertyName = "Size",
            PreviousValue = null,
            NewValue = complexValue
        };

        // Assert
        result.NewValue.Should().Be(complexValue);
        result.PreviousValue.Should().BeNull();
    }

    #endregion

    #region Cross-Model Tests

    [Fact]
    public void FormControlInfo_And_ReportControlInfo_ShouldHaveSimilarStructure()
    {
        // Arrange
        var formControl = new FormControlInfo
        {
            Name = "test",
            ControlType = "TextBox",
            ControlTypeId = 109,
            Section = "Detail",
            SectionId = 0,
            Left = 100,
            Top = 200,
            Width = 1000,
            Height = 500,
            Visible = true
        };

        var reportControl = new ReportControlInfo
        {
            Name = "test",
            ControlType = "TextBox",
            ControlTypeId = 109,
            Section = "Detail",
            SectionId = 0,
            Left = 100,
            Top = 200,
            Width = 1000,
            Height = 500,
            Visible = true
        };

        // Assert - Both should have same basic properties
        formControl.Name.Should().Be(reportControl.Name);
        formControl.ControlType.Should().Be(reportControl.ControlType);
        formControl.ControlTypeId.Should().Be(reportControl.ControlTypeId);
        formControl.Section.Should().Be(reportControl.Section);
        formControl.Left.Should().Be(reportControl.Left);
        formControl.Top.Should().Be(reportControl.Top);
        formControl.Width.Should().Be(reportControl.Width);
        formControl.Height.Should().Be(reportControl.Height);
        formControl.Visible.Should().Be(reportControl.Visible);
    }

    #endregion
}
