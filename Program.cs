using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var builder = WebApplication.CreateBuilder(args);

// Add controllers for API only (no Razor views needed)
builder.Services.AddControllers();

var app = builder.Build();

// Development exception page
if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}

app.UsePathBase("/LISImportSPCApi");

// Serve static files from wwwroot
app.UseStaticFiles();

app.UseRouting();
app.UseAuthorization();

// Map API controllers
app.MapControllers();

// Optional: fallback for non-API routes to redirect to your HTML page
//app.MapFallback(context =>
//{
//    context.Response.Redirect("/html/ImportExcel.html");
//    return System.Threading.Tasks.Task.CompletedTask;
//});

app.MapFallback(context =>
{
    context.Response.Redirect("/LISImportSPCApi/html/ImportExcel.html");
    return Task.CompletedTask;
});

app.Run();
