using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ToolSiteAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FilesController : ControllerBase
    {
        private readonly IHostingEnvironment env;

        public FilesController(IHostingEnvironment env)
        {
            this.env = env;
        }

        [HttpGet("Xlsx/{id}")]
        public IActionResult Xlsx(string id)
        {
            var filePath = Path.Combine(env.WebRootPath, "tmp", id);
            if (!System.IO.File.Exists(filePath))
                return NotFound();

            var memoryStream = new MemoryStream();
            using (var fs = System.IO.File.OpenRead(filePath))
                fs.CopyTo(memoryStream);
            System.IO.File.Delete(filePath);
            return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", id);
        }
    }
}