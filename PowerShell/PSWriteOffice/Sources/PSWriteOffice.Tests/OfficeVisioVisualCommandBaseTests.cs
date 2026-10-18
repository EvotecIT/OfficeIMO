using System.Management.Automation;
using ChartForgeX.VisualArtifacts;
using PSWriteOffice.Services.Visuals;
using Xunit;

namespace PSWriteOffice.Tests;

public sealed class OfficeVisioVisualCommandBaseTests
{
    [Fact]
    public void ScalarBytePipelineRejectsPayloadBeforeItExceedsTheEnvelopeLimit()
    {
        var command = new TestOfficeVisioVisualCommand();
        for (int index = 0; index < VisualArtifactInterchangeEnvelope.MaximumJsonUtf8Bytes; index++)
        {
            command.Buffer((byte)0);
        }

        PSArgumentOutOfRangeException exception = Assert.Throws<PSArgumentOutOfRangeException>(() => command.Buffer((byte)0));

        Assert.Contains("must not exceed", exception.Message);
        Assert.Equal(VisualArtifactInterchangeEnvelope.MaximumJsonUtf8Bytes, command.Complete().Length);
    }

    private sealed class TestOfficeVisioVisualCommand : OfficeVisioVisualCommandBase
    {
        public bool Buffer(object value) => BufferPipelineByte(value);

        public byte[] Complete() => CompletePipelineBytes()!;
    }
}
