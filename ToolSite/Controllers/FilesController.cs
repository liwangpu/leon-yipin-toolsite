using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using System.IO;

namespace ToolSite.Controllers
{

    public class FilesController : Controller
    {
        private readonly IHostingEnvironment env;

        public FilesController(IHostingEnvironment env)
        {
            this.env = env;
        }

        [HttpGet]
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