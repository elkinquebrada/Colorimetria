using ColorimetriaAPI.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new()
    {
        Title = "Colorimetría Token Corrector API",
        Version = "v1",
        Description = "API para corrección de tokens colorimétricos erróneos. " +
                      "Corrección 100% local — sin llamadas externas."
    });
});

builder.Services.AddScoped<ClaudeService>();

builder.Services.AddCors(options =>
{
    options.AddPolicy("LocalApp", policy =>
        policy.WithOrigins("http://localhost", "https://localhost")
              .AllowAnyHeader()
              .AllowAnyMethod());
});

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(c =>
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "Colorimetría API v1");
        c.RoutePrefix = string.Empty;
    });
}

app.UseCors("LocalApp");
app.UseAuthorization();
app.MapControllers();

app.Run();