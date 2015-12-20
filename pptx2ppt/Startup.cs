using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(pptx2ppt.Startup))]
namespace pptx2ppt
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
