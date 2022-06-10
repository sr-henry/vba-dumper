# vba-dumper
Simple application to extract/load all the vba project modules

```
//>dmp
usage: VBA Project Dumper [options]

Simple application to dump vba projects

optional arguments:
  -h, --help            show this help message and exit
  -v, --version         show program's version number and exit
  -d DUMP, --dump DUMP  dump VBA Modules
  -l LOAD, --load LOAD  load VBA Modules
```

You can edit the code outside the VBA Editor in a easy way.

Just need to extract, modify and load.

```
//>dmp --dump my_project.xlsm

//>dir
10/06/2022  18:16    <DIR>          my_project
10/06/2022  18:05           115.888 my_project.xlsm

//my_project>dir
10/06/2022  18:16    <DIR>          Class Modules
10/06/2022  18:16    <DIR>          Modules

//>dmp --load my_project.xlsm
```
