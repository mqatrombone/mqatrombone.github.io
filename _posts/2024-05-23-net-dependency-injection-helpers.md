---
layout: post
title: .NET Dependency Injection's ServiceCollection Extension Methods
date: 2024-05-23 20:00:00 -05:00
---
It is important when registering services to do the right thing. There's six
helper methods that will see typical use:

* TryAddSingleton
* TryAddScoped
* TryAddTransient
* AddSingleton
* AddScoped
* AddTransient

Is there only to be one implementation? Then you want a `TryAdd*`
method. Most of the time, you're only going to need a single implementation,
but when you want a collection of services implementing the same interface,
the `Add*` methods are there for you.

Now you need to choose between the available lifetimes:
* Singleton - a single instance will exist for the lifetime of the app. This
  doesn't mean that it initializes on registration, but that resolver will
  only generate one instance.
* Scoped - an instance will be reused for the life of the scope - typically
  this will be the life of an HttpRequest.
* Transient - every time the resolver will create a new instance. This is the
  simplest to implement due to the lack of state, but be careful.

.NET 8 adds keyed services, and there's also some other useful utilities for
quickly instantiating services at startup.

## References:

* [General .NET DI Overview](https://learn.microsoft.com/en-us/dotnet/core/extensions/dependency-injection)
* [.NET DI Guidelines](https://learn.microsoft.com/en-us/dotnet/core/extensions/dependency-injection-guidelines)
* [ASP.NET DI](https://learn.microsoft.com/en-us/aspnet/core/fundamentals/dependency-injection?view=aspnetcore-8.0)
