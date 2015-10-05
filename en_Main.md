![http://img232.imageshack.us/img232/922/monorunneruu0.png](http://img232.imageshack.us/img232/922/monorunneruu0.png)

# MONORunner: What is ... #

If you have already used a MONO application on MS Windows platform, you saw that to start the application you must specify thepath of MONO executable (mono.exe) and follow the path of application.
**MONORunner** executable try to lunch correctly the MONO application searching the right path for mono.exe.

# MONORunner: How use it ... #

I can use **MONORunner** to create setup package for MS Windows platform for MONO application.
Example:
If we have a MONO application named _myApplication.exe_ that we want distribute. We can create a setup package that contain _myApplication.exe_ and _monorunner.exe_ but the link that lunch the application must point to _monorunner.exe_ and have _myApplication.exe_ as parameter. **MONORunner**, after founded right position of mono.exe, use passed parameter to start the application.

[![](http://img505.imageshack.us/img505/5232/monopoweredbm0.png)](http://www.mono-project.com)