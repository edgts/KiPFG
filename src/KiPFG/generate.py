#!/usr/bin/env python3

import os
import re
import sys
import sexpdata
import subprocess
import argparse
import fitz

parser = argparse.ArgumentParser(
        prog="KiPFG",
        description="Generate production files for KiCad"
)

parser.add_argument(
    '-s',
    '--schematic-file',
    help="KiCad schematic file",
    type=str
)

parser.add_argument(
    '-p',
    '--pcb-file',
    help="KiCad pcb file",
    type=str
)

project_name = None

output_path_pdf = "PDF"
output_path_gerber = "FAB"
output_path_bom = "BOM"
output_path_3d = "3D"
output_path_rule_checks = "RCH"


def getProjectName():

    project_name = None

    for file in os.listdir(os.getcwd()):
        if file.endswith(".kicad_pro"):
            project_name = file.replace(".kicad_pro", "")
            break

    return project_name


def rreplace(s, old, new, occurrence):
    li = s.rsplit(old, occurrence)
    return new.join(li)


def sexParse(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        data = f.read()
        parsed_data = sexpdata.loads(data)
        return parsed_data


def getCopperLayerNames(data):

    layers = []

    for element in data:
        if element[0] == "layers":
            for layer in element:
                if layer[1].endswith(".Cu"):
                    layers.append(layer[1])

    return layers


def getRevision(data):
    revision = None

    for element in data:
        if element[0] == "title_blocK":
            for title_block_element in element:
                if title_block_element[0] == "rev":
                    revision = title_block_element[1]

    return revision


def exportPdfPcb(input_file, layers, revision):
    print("==== Export PCB PDF - Start ====")

    if not os.path.isdir(output_path_pdf):
        os.makedirs(output_path_pdf)

    pcb_basic_pdf_layers = [
        'F.Paste',
        'F.Mask',
        'F.Silkscreen',
        'F.Fab',
        'B.Paste',
        'B.Mask',
        'B.Silkscreen',
        'B.Fab',
    ]

    pcb_pdf_layers = layers + pcb_basic_pdf_layers

    # generate temporary single pdf files for each layer
    for layer in pcb_pdf_layers:

        output_path = os.path.join(os.getcwd(), output_path_pdf)
        output_filename = project_name + '_R' + revision + '_' + layer + '.pdf'
        output = os.path.join(output_path, output_filename)

        args = ['kicad-cli', 'pcb', 'export', 'pdf', '--ibt', '-l']
        args += [layer + ",Edge.Cuts", '--define-var', 'LAYER=' + layer,
                 '--output', output, input_file]

        subprocess.call(args)

    result = fitz.open()

    for layer in pcb_pdf_layers:

        output_path = os.path.join(os.getcwd(), output_path_pdf)
        output_filename = project_name + '_R' + revision + '_' + layer + '.pdf'
        pdf_file = os.path.join(output_path, output_filename)

        with fitz.open(pdf_file) as mfile:
            result.insert_pdf(mfile)

    pdf_save_path = os.path.join(os.getcwd(), output_path_pdf)
    pdf_save_path = os.path.join(pdf_save_path, project_name + '_R' + revision + '_PCB.pdf')
    result.save(pdf_save_path)

    for layer in pcb_pdf_layers:

        if layer == "F.Fab":
            continue

        if layer == "B.Fab":
            continue

        output_path = os.path.join(os.getcwd(), output_path_pdf)
        output_filename = project_name + '_R' + revision + '_' + layer + '.pdf'
        pdf_file = os.path.join(output_path, output_filename)
        os.remove(pdf_file)

    print("==== Export PCB PDF - End ====")


def exportDrc(input_file, revision):
    print("==== DRC - Start ====")

    if not os.path.isdir(output_path_rule_checks):
        os.makedirs(output_path_rule_checks)

    args = [
            "kicad-cli",
            "pcb",
            "drc",
            "--schematic-parity",
            "--format",
            "json",
            "--severity-warning",
            "--output",
            os.path.join(
                os.getcwd(),
                output_path_rule_checks,
                project_name + "_R" + revision + "_DRC.json"
            ),
            input_file
    ]

    subprocess.call(args)
    print("==== DRC - End ====")


def exportGerbers(input_file, copper_layer_list, revision):
    print("==== Export Gerber - Start ====")

    if not os.path.isdir(output_path_gerber):
        os.makedirs(output_path_gerber)

    pcb_basic_gerber_layers = [
        'F.Paste',
        'F.Mask',
        'F.Silkscreen',
        'F.Fab',
        'B.Paste',
        'B.Mask',
        'B.Silkscreen',
        'B.Fab',
        'Edge.Cuts'
    ]

    gerber_layers = copper_layer_list + pcb_basic_gerber_layers

    # Gerber files
    args = [
            "kicad-cli",
            "pcb",
            "export",
            "gerbers",
            "--layers",
            ",".join(gerber_layers),
            "--use-drill-file-origin",
            "--no-protel-ext",
            "--output",
            os.path.join(os.getcwd(), output_path_gerber),
            input_file
    ]

    subprocess.call(args)

    for layer in gerber_layers:
        filename = os.path.join(os.getcwd(), output_path_gerber, project_name + "-" + layer.replace(".", "_") + ".gbr")
        new_filename = rreplace(filename, project_name, project_name + "_R" + revision, 1)

        if os.path.isfile(filename):
            if os.path.isfile(new_filename):
                os.remove(new_filename)

            os.rename(filename, new_filename)

    # Drill files
    args = [
            "kicad-cli",
            "pcb",
            "export",
            "drill",
            "--drill-origin",
            "plot",
            "--excellon-separate-th",
            "--generate-map",
            "--map-format",
            "gerberx2",
            "--output",
            os.path.join(os.getcwd(), output_path_gerber),
            input_file
    ]

    subprocess.call(args)

    filenames = [
        os.path.join(os.getcwd(), output_path_gerber, project_name + "-NPTH.drl"),
        os.path.join(os.getcwd(), output_path_gerber, project_name + "-NPTH-drl_map.gbr"),
        os.path.join(os.getcwd(), output_path_gerber, project_name + "-PTH.drl"),
        os.path.join(os.getcwd(), output_path_gerber, project_name + "-PTH-drl_map.gbr"),
        os.path.join(os.getcwd(), output_path_gerber, project_name + "-job.gbrjob"),
    ]

    for filename in filenames:

        new_filename = rreplace(filename, project_name, project_name + "_R" + revision, 1)

        if os.path.isfile(filename):
            if os.path.isfile(new_filename):
                os.remove(new_filename)

            os.rename(filename, new_filename)

    print("==== Export Gerber - End ====")


def exportPickAndPlace(input_file, revision):
    print("==== Export Pick and place - Start ====")

    if not os.path.isdir(output_path_gerber):
        os.makedirs(output_path_gerber)

    args = [
        "kicad-cli",
        "pcb",
        "export",
        "pos",
        "--side",
        "front",
        "--format",
        "csv",
        "--units",
        "mm",
        "--use-drill-file-origin",
        "--output",
        os.path.join(os.getcwd(), output_path_gerber, project_name + "_R" + revision + "-top-pos.csv"),
        input_file
    ]

    subprocess.call(args)

    args = [
        "kicad-cli",
        "pcb",
        "export",
        "pos",
        "--side",
        "back",
        "--format",
        "csv",
        "--units",
        "mm",
        "--use-drill-file-origin",
        "--smd-only",
        "--output",
        os.path.join(os.getcwd(), output_path_gerber, project_name + "_R" + revision + "-bottom-pos.csv"),
        input_file
    ]

    subprocess.call(args)
    print("==== Export Pick and place - End ====")


def exportStep(input_file, revision):
    print("==== Export 3D - Start ====")

    if not os.path.isdir(output_path_3d):
        os.makedirs(output_path_3d)

    args = [
        "kicad-cli",
        "pcb",
        "export",
        "step",
        "--force",
        "--drill-origin",
        "--no-optimize-step",
        "--output",
        os.path.join(os.getcwd(), output_path_3d, project_name + "_R" + revision + "_3D.step"),
        input_file
    ]

    subprocess.call(args)
    print("==== Export 3D - End ====")

def getFilenameWithouthExtension(filename):
    return os.path.splitext(os.path.basename(filename))[0]

def exportPdfSch(input_file, revision):
    print("==== Export PDF SCH - Start ====")

    schematic_name = getFilenameWithouthExtension(input_file)

    if not os.path.isdir(output_path_pdf):
        os.makedirs(output_path_pdf)

    args = [
        "kicad-cli",
        "sch",
        "export",
        "pdf",
        "--output",
        os.path.join(os.getcwd(), output_path_pdf, schematic_name + "_R" + revision + "_SCH.pdf"),
        input_file,
    ]

    subprocess.call(args)


    print("==== Export PDF SCH - End ====")
    

# def exportBom(input_file, revision):


args = parser.parse_args()

if args.schematic_file:
    print("schematic file name passed via cli")
    schematic_file_name = re.search(r'\b\w+\.kicad_sch\b', args.schematic_file)

    if schematic_file_name:
        schematic_file_name = schematic_file_name.group().upper()
    else:
        print("Wrongly formatted schematic filename provided: '" + args.schematic_file + "'. Should be 'FILENAME.kicad_sch'.")
        sys.exit()

    if os.path.exists(os.path.join(os.getcwd(), )):
        schematic_file = schematic_file_name
    else:
        print("No kicad schematic under the name '" + schematic_file_name + "' could be found. Terminating")
        sys.exit()
else:
    print("project name automatically")
    project_name = getProjectName()


# if not project_name:
#     print("Error: No kicad project file could be found. Terminating.")
#     sys.exit()

print("Schematic name: '" + schematic_file + "'")


# input_file_pcb = project_name + ".kicad_pcb"
# input_file_sch = project_name + ".kicad_sch"

parsed_sch = sexParse(schematic_file)
# parsed_pcb = sexParse(input_file_pcb)

# pcb_copper_layer_list = getCopperLayerNames(parsed_pcb)

# pcb_revision = getRevision(parsed_pcb)
sch_revision = getRevision(parsed_sch)

# if sch_revision is pcb_revision:

#     revision = sch_revision

#     # PCB
      # exportPdfPcb(input_file_pcb, pcb_copper_layer_list, revision)
#     exportDrc(input_file_pcb, revision)
#     exportGerbers(input_file_pcb, pcb_copper_layer_list, revision)
#     exportPickAndPlace(input_file_pcb, revision)
#     exportStep(input_file_pcb, revision)

#     # SCH
exportPdfSch(schematic_file, sch_revision)

# else:
#     print("Schematic and PCB revision don't match.\nSCH Revision: " + sch_revision + "\nPCB Revision: " + pcb_revision + "\nAbort...")
