---
layout: post
title: .NET Framework + Microsoft.Extensions.DependencyInjection
date: 2024-05-21 20:00:00 -05:00
---
Say you're a developer, and you're stuck with a Frankenstein's monster of a
.NET Framework web application. No big deal, but you want to modernize a piece
of it. You're almost certainly want to involve `Microsoft.Extensions.DependencyInjection`,
the marvelous standard (modern) .NET Dependency Injection (DI) library. There's
good news: **it targets .NET Standard 2.0**, which means it is backwards compatible
with .NET Framework. The bad news? You're stuck manually wiring it up, because
ASP.NET MVC 5 doesn't do all the magic work that ASP.NET Core does for you.

One might search "ASP.NET MVC 5 Microsoft.Extensions.DependencyInjection" and see
if [someone](https://scottdorman.blog/2016/03/17/integrating-asp-net-core-dependency-injection-in-mvc-4/)
[has](https://learn.microsoft.com/en-us/answers/questions/807631/how-can-i-use-dependency-injection-in-asp-net-mvc)
[already](https://stackoverflow.com/questions/43311099/how-to-create-dependency-injection-for-asp-net-mvc-5)
[solved](https://stackoverflow.com/questions/50358349/using-microsoft-extension-dependencyinjection-in-asp-net-web-api2)
[this](https://stackoverflow.com/questions/55511586/mvc5-web-api-and-dependency-injection)
for you. You might get lucky and stumble across
[a gist from fowler](https://gist.github.com/davidfowl/563a602936426a18f67cd77088574e61)
or [a random engineering blog](https://www.eidias.com/blog/2024/1/3/implement-microsoft-di-container-in-aspnet-mvc-5).
and find your solution. This is an attempt to further codify this knowledge
and make it easier to discover.

<!--more-->

## What is Microsoft.Extensions.DependencyInjection?

If you're going to use Microsoft's tooling, it is important to understand how
all of it fits together. There's two parts involved:
* `Microsoft.Extensions.DependencyInjection.Abstractions` - the interfaces
* `Microsoft.Extensions.DependencyInjection` - the implementation
Basically, the infrastructure/startup of your application needs the implementation
and any helper functions living in other projects needs the interfaces.

`M.E.DI` is a (relatively simple) Inversion of Control (IoC) container system.
Most developer interaction with `M.E.DI` in modern .NET involves preparing
services and using them. That's the well-documented part and I will refer you to
[learn.microsoft.com](https://learn.microsoft.com/en-us/aspnet/core/fundamentals/dependency-injection)
for more information. The ServiceCollection and its extension methods allow
developers to register services with different lifetimes. After the services
are added to the ServiceCollection, the developer can use the collection to
build a ServiceProvider used by the other part of the DI equation.

The other major part is the Resolver, which more accurately called the
Container. This is the piece that wires up the registered services and provides
implementations when called upon. This is where it gets tricky for .NET Framework,
as the app developer will likely need to provide one. Microsoft also left a bit
of a trap here, as the different service lifetimes can cause us fits.

## What do I do?

Start at the beginning. Build a function to configure your services.

```csharp
public void ConfigureServices
{
    var services = new ServiceCollection();

    /* register your services here */
    services.AddTransient<IMyAwesomeService>();
    services.AddLogging();
    services.AddHttpClient();
    services.MyExtensionToRegisterServices();
}
```

Great, now you have a serviceCollection, but it isn't useful, because
you haven't created a custom resolver. You'll want a resolver that
will create a custom scope for each `HttpRequest`.

```csharp
public sealed class ScopedDependencyResolver : IDependencyResolver
{
    private readonly IServiceProvider _serviceProviderRoot;

    public ScopedDependencyResolver(IServiceProvider serviceProviderRoot)
    {
        _serviceProviderRoot = serviceProviderRoot;
    }

    public object GetService(Type serviceType)
    {
        return ResolveServiceProvider().GetService(serviceType);
    }

    public IEnumerable<object> GetServices(Type serviceType)
    {
        return ResolveServiceProvider().GetServices(serviceType);
    }
    
    private IServiceProvider ResolveServiceProvider()
    {
        if (HttpContext.Current == null)
        {
            return _serviceProviderRoot;
        }

        var scope = HttpContext.Current.Items[typeof(IServiceScope)] as IServiceScope;
        if (scope == null)
        {
            scope = _serviceProviderRoot.CreateScope();
            HttpContext.Current.Items[typeof(IServiceScope)] = scope;
        }

        reutrn scope.ServiceProvider;
    }
}
```
This is slightly risky - we won't dispose of scopes automatically,
so we can add an HttpModule to do it for us.
```csharp
public class ScopedDependencyHttpModule : IHttpModule
{
    public void Init(HttpApplication context)
    {
        context.EndRequest += Context_EndRequest;
    }

    private void Context_EndRequest(object sender, EventArgs e)
    {
        var context = ((HttpApplication)sender).Context;
        if (context.Items[typeof(IServiceScope)] is IServiceScope)
        {
            scope.Dispose();
            contexxt.Items.Remove(typeof(IServiceScope));
        }
    }

    public void Dispose() { }
}
```
Now that is complete, we'll need to ensure it gets loaded early in our
startup.

```csharp
[assembly: PreApplicationStartMethod(typeOf(ApplicationClass), "InitModules")]
public class ApplicationClass : HttpApplication
{
    public static void InitModules()
    {
        RegisterModule(typeof(ScopedDependencyHttpModule));
    }
}
```
And now we can finish up our `ConfigureServices` function
```csharp
public void ConfigureServices()
{
    var services = new ServiceCollection();

    /* register your services here */
    services.AddTransient<IMyAwesomeService>();
    services.AddLogging();
    services.AddHttpClient();
    services.MyExtensionToRegisterServices();

    /* New stuff */
    var rootProvider = services.BuildServiceProvider(true);
    var resolver = new ScopedDependencyResolver(rootProvider);
    DependencyResolver.SetResolver(resolver);
}
```

\<Bluey\>Hooray!\</Bluey\>

However, now we need to grab our Controllers. The easiest way to do
that is to add a `AddControllers` extension class to `IServiceCollection`.
It might look something like this:
```csharp
public static void AddControllers(this IServiceCollection services)
{
    var controllers = typeof(MyClass).Assembly.GetExportedTypes()
                        .Where(t => !t.IsAbstract && !t.IsGenericTypeDefinition)
                        .Where(t => typeOf(IController).IsAssignableFrom(t)
                                 || t.Name.EndsWith("Controller", StringComparison.OrdinalIgnoreCase);

    foreach(var c in controllers)
    {
        services.AddTransient(c);
    }

    return services;
}
```
