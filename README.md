# Connect To Program
Originally this code was built into the dotnet framework and was a simple
operation to perform the connection to a program based on the ProgID...

When the upgrade to dotnet 5 meant that the dotnet core base was used to replace
the dotnet framework and the code bases were consolidated, platform specific
features such as the marshall tool needed were no longer included. What was
simple code now required loading _OLE32_ dlls to get the same effect.

This library deals with the details and reduces it back to 1 line of code.

C# Code to use this library.

``` C#

(TopAPIObject?) ConnectToProgram.GetActiveCOMConnection.GetActiveConnection("ProgID...");

```
The casting operation is important as you will need to load the correct dll and
types so that the cast from the running object table of windows will work
correctly. It make take some effort to find the right API object, but it is most
likely the highest level object in the tree.

The question mark on the cast is also important as if no running instance of the
program is found it will return null.

The ProgID can be googled if you are not sure what it is. But it ties to the
windows registry and is usually a meaningful name, not a GUID. Note, you may
need to update the ProgID with each annual release of the program you are
accessing depending on how the registry values are setup on installation of the
program in question.

Refer to the Program API documentation to find what the top level object is and
from there you can dig into the functionality available for the running program.
The DLL file will need to be referenced in your project and the object used for
the casting operation. Dynamic values are used inside this program so there
should be a passthough without needing to specify types at this level of the
library.