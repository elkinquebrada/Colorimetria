using ColorimetriaAPI.Middleware;
using ColorimetriaAPI.Services;

var builder = WebApplication.CreateBuilder(args);

// ── Servicios ──────────────────────────────────────────────────────────────

builder.Services.AddControllers();

// Servicio colorimétrico 100% local — sin llamadas externas
builder.Services.AddScoped<ColorimetryService>();

// Swagger solo en desarrollo (nunca exponer en producción)
if (builder.Environment.IsDevelopment())
{
    builder.Services.AddEndpointsApiExplorer();
    builder.Services.AddSwaggerGen(c =>
    {
        c.SwaggerDoc("v1", new() { Title = "ColorimetríaAPI", Version = "v1" });

        // Añadir campo de API Key en Swagger para pruebas locales
        c.AddSecurityDefinition("ApiKey", new Microsoft.OpenApi.Models.OpenApiSecurityScheme
        {
            Name = "X-Api-Key",
            In = Microsoft.OpenApi.Models.ParameterLocation.Header,
            Type = Microsoft.OpenApi.Models.SecuritySchemeType.ApiKey,
            Description = "Clave de autenticación. Header: X-Api-Key"
        });
        c.AddSecurityRequirement(new Microsoft.OpenApi.Models.OpenApiSecurityRequirement
        {
            {
                new Microsoft.OpenApi.Models.OpenApiSecurityScheme
                {
                    Reference = new Microsoft.OpenApi.Models.OpenApiReference
                    {
                        Type = Microsoft.OpenApi.Models.ReferenceType.SecurityScheme,
                        Id   = "ApiKey"
                    }
                },
                Array.Empty<string>()
            }
        });
    });
}

// ── Seguridad: HTTPS y HSTS ────────────────────────────────────────────────

builder.Services.AddHsts(options =>
{
    options.MaxAge = TimeSpan.FromDays(365);
    options.IncludeSubDomains = true;
});

// ── Construir app ──────────────────────────────────────────────────────────

var app = builder.Build();

// Forzar HTTPS en todos los entornos
app.UseHttpsRedirection();

// HSTS solo en producción
if (!app.Environment.IsDevelopment())
    app.UseHsts();

// Swagger solo en desarrollo
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(c =>
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "ColorimetríaAPI v1");
        c.RoutePrefix = string.Empty; // Swagger en la raíz http://localhost:5000
    });
}

// Autenticación por API Key (antes de los controllers)
app.UseMiddleware<ApiKeyMiddleware>();

app.UseAuthorization();
app.MapControllers();

app.Run();