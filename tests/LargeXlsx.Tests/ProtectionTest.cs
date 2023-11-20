using System.IO;
using FluentAssertions;
using NUnit.Framework;
using OfficeOpenXml;

namespace LargeXlsx.Tests;

[TestFixture]
public static class ProtectionTest
{
    [Test]
    public static void Locked()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1").BeginRow()
                .Write("Test", XlsxStyle.Default.With(XlsxProtection.Default.WithLocked(false)));
        using (var package = new ExcelPackage(stream))
            package.Workbook.Worksheets[0].Cells["A1"].Style.Locked.Should().BeFalse();
    }

    [Test]
    public static void HiddenFormula()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1").BeginRow()
                .Write("Test", XlsxStyle.Default.With(XlsxProtection.Default.WithHiddenFormula(true)));
        using (var package = new ExcelPackage(stream))
            package.Workbook.Worksheets[0].Cells["A1"].Style.Hidden.Should().BeTrue();
    }
}
