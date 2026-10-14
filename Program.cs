using Microsoft.Agents.Hosting.AspNetCore;
using Serilog;
using Microsoft.Agents.Storage;
using Microsoft.AspNetCore.Http.Features;
using NxSphereSupportAgent.Agents;
using NxSphereSupportAgent.Investigation;
using NxSphereSupportAgent.Pages;

var builder = WebApplication.CreateBuilder(args);

// ---------------------------------------------------------------------------
// Logging
// ---------------------------------------------------------------------------
// Console alone is not enough: when an investigation fails, the evidence has to
// outlive the terminal. {Exception} appears in the file template only, so full
// stack traces are written to disk and never anywhere they could reach Teams.
//
// logs/ holds case file names and subject lines (already logged today), so treat
// it as customer-adjacent. Email bodies are never logged.
builder.Host.UseSerilog((context, services, configuration) => configuration
    .MinimumLevel.Information()
    .MinimumLevel.Override("NxSphereSupportAgent", Serilog.Events.LogEventLevel.Debug)
    // Verbose SDK logging is fine on a console that scrolls away; it is a
    // different proposition once it is retained on disk for two weeks.
    .MinimumLevel.Override("Microsoft", Serilog.Events.LogEventLevel.Warning)
    .Enrich.FromLogContext()
    .WriteTo.Console()
    .WriteTo.File(
        Path.Combine(context.HostingEnvironment.ContentRootPath, "logs", "support-agent-.log"),
        rollingInterval: RollingInterval.Day,
        retainedFileCountLimit: 14,
        outputTemplate:
            "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {SourceContext} {Message:lj}{NewLine}{Exception}"));

// ---------------------------------------------------------------------------
// Teams personal-chat agent (Microsoft 365 Agents SDK)
// ---------------------------------------------------------------------------
// AddAgentDefaults registers the HttpClient factory, controllers and core SDK
// services. MemoryStorage keeps turn state in process - fine for a POC, and it
// means nothing is persisted between runs.
builder.AddAgentDefaults();
builder.Services.AddSingleton<IStorage, MemoryStorage>();
builder.AddAgentApplicationOptions();
builder.AddAgent<NxSphereSupportAgentBot>();

// ---------------------------------------------------------------------------
// Case investigation (Claude, via the `claude` CLI + local knowledge base)
// ---------------------------------------------------------------------------
// The CLI authenticates itself against a Claude Pro/Max/Team subscription
// (`claude setup-token` / `CLAUDE_CODE_OAUTH_TOKEN`), not a Console API key,
// so there is no secret for this app to hold via dotnet user-secrets.
builder.Services.Configure<InvestigationOptions>(
    builder.Configuration.GetSection(InvestigationOptions.SectionName));
builder.Services.Configure<ClaudeCliOptions>(
    builder.Configuration.GetSection(ClaudeCliOptions.SectionName));

builder.Services.AddHttpClient();

// Stateless once constructed, so one instance is shared for the app's
// lifetime rather than built per request - same reasoning as the direct-API
// version's AnthropicClient singleton.
builder.Services.AddSingleton<ClaudeCliClient>();

// Loaded lazily on first use and then reused, so the ~330 KB knowledge base is
// read from disk once, not on every case.
builder.Services.AddSingleton<KnowledgeBase>();
builder.Services.AddSingleton<OperatingInstructions>();
builder.Services.AddSingleton<EmailParser>();
builder.Services.AddSingleton<CaseFileDownloader>();
builder.Services.AddSingleton<ClaudeInvestigator>();
builder.Services.AddSingleton<InlineImageStore>();

// ---------------------------------------------------------------------------
// Razor Pages upload page, retained from the tab phase
// ---------------------------------------------------------------------------
// No longer listed in the Teams manifest, so Teams never renders it. Kept as a
// local test harness: its .eml validation is what Milestone 2 will reuse.
builder.Services.AddRazorPages();

builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = IndexModel.MaxUploadBytes;
});

builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = IndexModel.MaxUploadBytes;
});

builder.Services.AddAntiforgery(options =>
{
    // Left from the tab phase: a cross-site iframe never returns a Strict cookie.
    options.Cookie.SameSite = SameSiteMode.None;
    options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
    options.Cookie.Extensions.Add("Partitioned");
});

var app = builder.Build();

// Inbound Azure Bot Service tokens are NOT yet validated - see the runbook. The
// agent endpoint is mapped anonymously, which the SDK permits for development
// only. Fail loudly rather than let that reach a real environment.
if (!app.Environment.IsDevelopment())
{
    throw new InvalidOperationException(
        "Inbound bot token validation is not configured in this POC phase, so the app "
        + "must only run with ASPNETCORE_ENVIRONMENT=Development. Configure "
        + "AddAgentAuthorization/UseAgents before running anywhere else.");
}

app.UseHttpsRedirection();

app.Use(async (context, next) =>
{
    context.Response.OnStarting(() =>
    {
        // Left from the tab phase; harmless for the bot endpoint.
        context.Response.Headers.Remove("X-Frame-Options");
        context.Response.Headers["Content-Security-Policy"] =
            "frame-ancestors teams.microsoft.com *.teams.microsoft.com *.skype.com *.office.com";
        context.Response.Headers["X-Content-Type-Options"] = "nosniff";
        return Task.CompletedTask;
    });

    await next();
});

// Puts a readable sentence on empty error responses from the upload page. Skipped
// for /api/* so a bot error reaches Azure Bot Service as-is instead of as prose.
app.UseWhen(
    context => !context.Request.Path.StartsWithSegments("/api"),
    branch => branch.UseStatusCodePages(async context =>
    {
        var response = context.HttpContext.Response;
        response.ContentType = "text/plain; charset=utf-8";

        var message = response.StatusCode switch
        {
            StatusCodes.Status413PayloadTooLarge =>
                $"Upload rejected: the file exceeds the {IndexModel.MaxUploadDisplay} limit for this POC.",
            StatusCodes.Status400BadRequest =>
                "Upload rejected: the request could not be validated. Reload the page and try again. "
                + $"If the file is larger than {IndexModel.MaxUploadDisplay} it will always be rejected.",
            StatusCodes.Status404NotFound => "Not found.",
            _ => $"Request failed with status {response.StatusCode}."
        };

        await response.WriteAsync(message);
    }));

// Serves wwwroot only. UseDirectoryBrowser is never called, so there is no listing.
app.UseStaticFiles();
app.UseRouting();

app.MapRazorPages();

// The messaging endpoint Azure Bot Service posts activities to.
// First argument false = allow anonymous; see the guard above.
app.MapAgentEndpoints(false, "/api/messages");

// Resolve the knowledge base and operating instructions now rather than on the
// first case. Both are Lazy with ExecutionAndPublication, which caches a thrown
// exception for the process lifetime - so a bad path would otherwise fail every
// case until someone restarted, long after the console had scrolled away.
await StartupChecks.RunAsync(app);

app.Run();
