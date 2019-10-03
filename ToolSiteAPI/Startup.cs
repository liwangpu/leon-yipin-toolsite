using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System.IO;

namespace ToolSiteAPI
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddCors(options =>
            {
                options.AddPolicy("CorsPolicy",
                    builder => builder.AllowAnyOrigin()
                    .AllowAnyMethod()
                    .AllowAnyHeader());
            });
            services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_2);
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            app.UseStaticFiles();
            app.UseCors("CorsPolicy");
            if (!Directory.Exists(env.WebRootPath))
            {
                if (string.IsNullOrEmpty(env.WebRootPath))
                {
                    env.WebRootPath = Path.Combine(env.ContentRootPath, "wwwroot");
                }
                Directory.CreateDirectory(env.WebRootPath);
            }
            var tmpFolder = Path.Combine(env.WebRootPath, "tmp");
            if (!Directory.Exists(tmpFolder))
                Directory.CreateDirectory(tmpFolder);

            var cacheFolder = Path.Combine(env.WebRootPath, "cache");
            if (!Directory.Exists(cacheFolder))
                Directory.CreateDirectory(cacheFolder);
            app.UseMvc();
        }
    }
}
