# LECmd

## Command Line Interface

    LECmd version 1.4.0.0
    
    Author: Eric Zimmerman (saericzimmerman@gmail.com)
    https://github.com/EricZimmerman/LECmd
    
            d               Directory to recursively process. Either this or -f is required
            f               File to process. Either this or -d is required
            q               Only show the filename being processed vs all output. Useful to speed up exporting to json and/or csv. Default is FALSE
    
            r               Only process lnk files pointing to removable drives. Default is FALSE
            all             Process all files in directory vs. only files matching *.lnk. Default is FALSE
    
            csv             Directory to save CSV formatted results to. Be sure to include the full path in double quotes
            csvf            File name to save CSV formatted results to. When present, overrides default name
    
            xml             Directory to save XML formatted results to. Be sure to include the full path in double quotes
            html            Directory to save xhtml formatted results to. Be sure to include the full path in double quotes
            json            Directory to save json representation to. Use --pretty for a more human readable layout
            pretty          When exporting to json, use a more human readable layout. Default is FALSE
    
            nid             Suppress Target ID list details from being displayed. Default is FALSE
            neb             Suppress Extra blocks information from being displayed. Default is FALSE
    
            dt              The custom date/time format to use when displaying time stamps. See https://goo.gl/CNVq0k for options. Default is: yyyy-MM-dd HH:mm:ss
            mp              Display higher precision for time stamps. Default is FALSE
    
    Examples: LECmd.exe -f "C:\Temp\foobar.lnk"
              LECmd.exe -f "C:\Temp\somelink.lnk" --json "D:\jsonOutput" --pretty
              LECmd.exe -d "C:\Temp" --csv "c:\temp" --html c:\temp --xml c:\temp\xml -q
              LECmd.exe -f "C:\Temp\some other link.lnk" --nid --neb
              LECmd.exe -d "C:\Temp" --all
    
              Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes    

## Documentation

LNK Explorer Command Line edition!

[Introducing LECmd!](https://binaryforay.blogspot.com/2016/02/introducing-lecmd.html)

[LECmd v0.6.0.0 released!](https://binaryforay.blogspot.com/2016/02/lecmd-v0600-released.html)

[PECmd, LECmd, and JLECmd updated!](https://binaryforay.blogspot.com/2016/03/pecmd-lecmd-and-jlecmd-updated.html)

[LECmd and JLECmd updated](https://binaryforay.blogspot.com/2016/04/lecmd-and-jlecmd-updated.html)

# Download Eric Zimmerman's Tools

All of Eric Zimmerman's tools can be downloaded [here](https://ericzimmerman.github.io/#!index.md). Use the [Get-ZimmermanTools](https://f001.backblazeb2.com/file/EricZimmermanTools/Get-ZimmermanTools.zip) PowerShell script to automate the download and updating of the EZ Tools suite. Additionally, you can automate each of these tools using [KAPE](https://www.kroll.com/en/services/cyber-risk/incident-response-litigation-support/kroll-artifact-parser-extractor-kape)!

# Special Thanks

Open Source Development funding and support provided by the following contributors: [SANS Institute](http://sans.org/) and [SANS DFIR](http://dfir.sans.org/).
