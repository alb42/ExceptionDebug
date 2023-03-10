# ExceptionDebug
A Delphi unit filling the stacktrace string in exceptions

The unit is created from the jedi jcl library https://github.com/project-jedi/jcl
but you don't need the whole jcl to use it all the necessary sources are put
in one single unit.

The library is released to the public under the terms of the Mozilla Public License (MPL) and as such can be freely used in both freeware/shareware, opensource and commercial projects. The entire JEDI Code Library is distributed under the terms of the Mozilla Public License (MPL).

## How to use

Add the unit to your project.
e.g.
```
try
  raise Exception.Create('New exception');
except
  on E:Exception do
    ShowMessage('Exception ' + E.Message + #13#10' Stacktrace: '#13#10 + E.StackTrace);
end;
```
Compile with debug symbols and you get a nice stacktrace where the exception happend.
