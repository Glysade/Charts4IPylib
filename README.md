# Charts4IPylib
IronPython library for configuring ChemCharts4 visuals

To use this module, configure the following in the ScriptSync Preferences:

![image](https://user-images.githubusercontent.com/46694342/198642238-f4fba0e8-076a-48c8-a0fd-0b7f1756b9c1.png)

Use these values to Copy+Paste:
```
Name: Charts4IPylib
Url: https://github.com/Glysade/Charts4IPylib
Branch: master
```

Import the Charts4IPylib using the following IronPython:

```
import sys
import __builtin__
from System.IO import Path
sys.path.append(Path.Combine(Path.GetTempPath(),'ScriptSync','Charts4IPylib'))
__builtin__.Document = Document
__builtin__.Application = Application
import Charts4IPylib
```
