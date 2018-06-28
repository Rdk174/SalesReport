using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(SalesReport.Startup))]
namespace SalesReport
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            //ConfigureAuth(app);
        }
    }
}
