using System.CommandLine;
using CapabilityModule.Office.Cli.Commands;
using Npgsql;

namespace CapabilityModule.Office.Cli.Tests;

/// <summary>
/// Exercises the `index build` command's exit-code contract (fixes #8). Per-file
/// errors (a corrupted document, etc.) must be reported in the JSON summary's
/// filesWithErrors count, not turned into a non-zero exit — WebApi and MCP
/// callers (<c>CliRunner</c> in each layer) treat any non-zero exit as a hard
/// failure and discard the JSON body entirely, so a single bad document would
/// otherwise present as "reindex failed" with no summary, even though every
/// other document indexed fine.
/// </summary>
[Collection("Postgres")]
public sealed class IndexCommandTests : IAsyncLifetime
{
    private readonly PostgresFixture _postgres;
    private readonly string _root;
    private readonly RootCommand _rootCmd;

    public IndexCommandTests(PostgresFixture postgres)
    {
        _postgres = postgres;
        _root = Path.Combine(Path.GetTempPath(), "office-cli-index-command-tests-" + Guid.NewGuid());
        Directory.CreateDirectory(_root);
        _rootCmd = new RootCommand { new IndexCommand().Command() };
    }

    public async Task InitializeAsync()
    {
        await using var dataSource = NpgsqlDataSource.Create(_postgres.ConnectionString);
        await using var conn = await dataSource.OpenConnectionAsync();
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "TRUNCATE documents, chunks RESTART IDENTITY CASCADE";
        await cmd.ExecuteNonQueryAsync();
    }

    public Task DisposeAsync()
    {
        if (Directory.Exists(_root))
            Directory.Delete(_root, recursive: true);
        return Task.CompletedTask;
    }

    [Fact]
    public async Task Build_FileWithExtractionError_StillExitsZero()
    {
        DocxEngine.Create(Path.Combine(_root, "good.docx"), "Good", "This one extracts fine.");
        // A .docx extension the OpenXml extractor cannot actually parse — a
        // per-file extraction error, not an unsupported-format skip.
        File.WriteAllText(Path.Combine(_root, "bad.docx"), "not a real docx — corrupted");

        Environment.SetEnvironmentVariable("OFFICE_DB_CONNECTION", _postgres.ConnectionString);
        try
        {
            var exitCode = await _rootCmd.InvokeAsync(
                new[] { "index", "build", ".", "--root", _root, "--embed", "false" });

            Assert.Equal(0, exitCode);
        }
        finally
        {
            Environment.SetEnvironmentVariable("OFFICE_DB_CONNECTION", null);
        }
    }
}
