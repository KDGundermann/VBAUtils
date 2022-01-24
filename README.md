VBAUtils
========

This is a collection of several util libraries I have created for VBA/VB6 in the last 20 years.

The libaries are structured in different subfolders.
As VBA has no native namespacing, I am using a simulated namespacing to keep a cleaner structure.
You will find classes call XX__Namespace and XX__Facade/Factory, e.g:

- OS__Namespace
- OS__Facade
- OS_Time

So in your code you can call functions from the OS_Time Library e.g. by using
`OS.Time.GetLocalTime()`

As there are several classes enhancing the builtin classes in VBA I have prefixed
the names with `°`  (YES, this is a valid character in VBA!) 

So you will have functions in:

- Strings  -> the VBA.Strings library
- °Strings -> the VBAUtils.°Strings library

