using System.CommandLine;

namespace Ox;

public static class GlobalOptions
{
    public static readonly Option<bool> DryRun = new("--dry-run", "-n") { Description = "Preview changes without modifying files", Recursive = true };
    public static readonly Option<bool> Force = new("--force", "-y") { Description = "Skip confirmations and force past lock files", Recursive = true };
}
