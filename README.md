# DXL.vbs

## Overview

DXL(Domino XML Language) export/import tool in VBScript.

You can export your HCL Notes database design into DXL(XML) file, so that you can manage design changes in text( using GitHub for example ).

You can also import your result into HCL Notes database(NSF) again from any git log history(still not working).


## Pre-requisite

- Windows

- HCL Notes setup

- git CLI installed and setup in your Windows

  - https://git-scm.com/download/win
  
  - `> git config --global user.name "Your Name"`
  
  - `> git config --global user.email "Your Email"`

- git account ( in GitHub, for example )


## How to export

- Open CMD(Command Line Prompt) in your Windows

- Navigate CMD in your working folder:

  - `> mkdir c:\tmp`
  
  - `> cd c:\tmp`

- Git clone this repository(only first once):

  - `> git clone https://github.com/dotnsf/dxl.vbs`

- Navigate CMD in your `dxl.vbs` folder:

  - `> cd dxl.vbs`
  - `> rmdir .git`

- Run `dxl_export.vbs` with your target Notes DB file path( `dev/sample.nsf`, for example ):

  - `> c:\Windows\SysWOW64\CScript //nologo dxl_export.vbs dev/sample.nsf`
  
  - **(Caution)** You need to specify this **c:\Windows\SysWOW64** full path in 64bit Windows(otherwise 32bit CScript.exe will be launched).
  
- If you are asked your Notes' password, enter it

- If succeeded, created result filepath would be shown:

  - `dev_sample.nsf/sample.xml`, for example

- Go to exported foloder:

  - `> cd dev_sample.nsf`

- If you want to manage your changes with Git(GitHub, for example), initialize this folder as Git repository:

  - `> git init`
  - `> git branch -M main`

- You can see text-exported design element with text-editor in result file

- You can commit this resulted folder(`dev_sample.nsf`, in this case) into Git:

  - `> git add .`
  
  - `> git commit -m 'some changes'`
  
  - `> cd ..`

- Edit your DB with Domino Designer

- Run `dxl_export.vbs` again with your target Notes DB, and commit changes:

  - `> c:\Windows\SysWOW64\CScript //nologo dxl_export.vbs dev/sample.nsf`
  
  - `> cd dev_sample.nsf`

  - `> git add .`
  
  - `> git commit -m 'new changes'`

- Then you can check design differences with `git diff` command:

  - `> git log`
  
    - find Commit IDs
  
  - `> git diff (Commit ID 1) (Commit ID 2)`


## Customized export

- If you want your documents also included in this management:

  - Edit `dxl_export.vbs` as `nc.SelectDocuments = **True**`


## How to import(still not working)

- Run `dxl_import.vbs` with your target XML file **full** path( `dev_sample.nsf\sample.xml`, for example ):

  - `> c:\Windows\SysWOW64\CScript //nologo dxl_import.vbs C:\Users\yourname\dev_sample.nsf\sample.xml`
  
  - **(Caution)** You need to specify this **c:\Windows\SysWOW64** full path in 64bit Windows(otherwise 32bit CScript.exe will be launched).
  
- If you are asked your Notes' password, enter it

- If succeeded, created result filepath would be shown:

  - `dev_sample.nsf/sample.nsf`, for example


## Licensing

This code is licensed under MIT.


## Copyright

2023 K.Kimura @ Juge.Me all rights reserved.

