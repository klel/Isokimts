using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(kimts.Startup))]
namespace kimts
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
