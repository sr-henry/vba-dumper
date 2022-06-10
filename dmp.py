import os
import argparse
import win32com.client as w32

VBTypes = {
    1 : "Modules",
    2 : "Class Modules"
}

def dump(wb_name: str):
    project_path = os.path.join(os.getcwd(), os.path.splitext(wb_name)[0])
    wb_path = os.path.join(os.getcwd(), wb_name)
    wb = w32.GetObject(wb_path)
    try:     
        os.mkdir(project_path)
        os.mkdir(os.path.join(project_path, "Modules"))
        os.mkdir(os.path.join(project_path, "Class Modules"))

        for m in wb.VBProject.VBComponents:
            if m.type in VBTypes.keys():
                path = os.path.join(project_path, VBTypes[m.type], m.name)
                m.Export(path)
                print(f"-{path}")
        
        wb = None
    except Exception as err:
            print(err)

def load(wb_name: str):
    project_path = os.path.join(os.getcwd(), os.path.splitext(wb_name)[0])
    wb_path = os.path.join(os.getcwd(), wb_name)
    wb = w32.GetObject(wb_path)
    for mdir in os.listdir(project_path):
        cpath = os.path.join(project_path, mdir)
        for f in os.listdir(cpath):
            m = wb.VBProject.VBComponents(f.split(".")[0])
            wb.VBProject.VBComponents.Remove(m)
            path = os.path.join(cpath, f)
            wb.VBProject.VBComponents.Import(path)
            print(f"+{path}")
    wb = None

if __name__ == "__main__":

    parser = argparse.ArgumentParser(
        prog="VBA Project Dumper",
        description="Simple application to dump vba projects",
        usage="%(prog)s [options]"
    )

    parser.add_argument("-v", "--version", action="version", version="1.0.0")
    parser.add_argument("-d", "--dump", help="dump VBA Modules", type=str)
    parser.add_argument("-l", "--load", help="load VBA Modules", type=str)
    
    parse_args = parser.parse_args()
    
    if parse_args.dump:
        dump(parse_args.dump)
    elif parse_args.load:
        load(parse_args.load)
    else:
        parser.print_help()