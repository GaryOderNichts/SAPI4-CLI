# SAPI4 CLI
Use Microsoft Sam & friends from the commandline.  
Based on [https://github.com/TETYYS/SAPI4](https://github.com/TETYYS/SAPI4).

### Example
`sapi4.exe -v "Adult Female #1, American English (TruVoice)" -s 100 -p 100 -t "Hello!" -o test.wav`  
Will create a file called `test.wav`.

### Usage
```
Usage: SAPI4 [options...]
Options:
    -v, --voice            voice                   (Required)
    -s, --speed            speed                   (Required)
    -p, --pitch            pitch                   (Required)
    -t, --text             text                    (Required)
    -o, --output           output file             (Required)
    -h, --help             Shows this page
```

### Build
1. Install Microsoft Speech SDK 4.0 (SAPI4SDK.exe). Can be found in the original SAPI4 repo.
2. Run `vcvars32.bat` of your installation of Visual Studio. I used `Microsoft Visual Studio 14.0`. Newer versions should work too.
3. Run `build.bat` to compile.

### Credits
TETYYS for the [original SAPI4 repo](https://github.com/TETYYS/SAPI4)  
Jesse Laning for the [argparser](https://github.com/jamolnng/argparse)  