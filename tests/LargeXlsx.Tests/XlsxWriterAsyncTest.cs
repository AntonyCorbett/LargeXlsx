using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FluentAssertions;
using NUnit.Framework;

namespace LargeXlsx.Tests;

[TestFixture]
public static class XlsxWriterAsyncTest
{
    [Test]
    public static void DisposeWriterTwice()
    {
        using var stream = new MemoryStream();
        var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet 1").BeginRowAsync().Write("Hello World!");
        xlsxWriter.Dispose();

        var act = () => xlsxWriter.Dispose();
        act.Should().NotThrow();
    }
}