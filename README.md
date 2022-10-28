# Charts4IPylib
IronPython library for configuring ChemCharts4 visuals

To use this module, configure the following in the ScriptSync Preferences:

![image](https://user-images.githubusercontent.com/46694342/198631923-d7cccca5-5074-49e3-9879-7f07a8797fe6.png)

Import the Charts4IPylib using the following IronPython:

```
import sys
import __builtin__
from System.IO import Path
sys.path.append(Path.Combine(Path.GetTempPath(),'ScriptSync','ConfigureCharts4'))
__builtin__.Document = Document
__builtin__.Application = Application
import Charts4IPylib
```
