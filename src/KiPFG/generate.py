#!/usr/bin/env python3

import os
import re
import sys
import subprocess
import fitz

from project_information import ProjectInformation

output_path_pdf = "PDF"
output_path_gerber = "FAB"
output_path_bom = "BOM"
output_path_3d = "3D"
output_path_rule_checks = "RCH"


class GeneratorError(Exception):
    pass


def rreplace(s, old, new, occurrence):
    li = s.rsplit(old, occurrence)
    return new.join(li)


def exportPdfPcb(input_file, layers, revision):
    print("* Generating pcb pdf files...", end="")

    pcb_name = getFilenameWithouthExtension(input_file)

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
        output_filename = pcb_name + '_R' + revision + '_' + layer + '.pdf'
        output = os.path.join(output_path, output_filename)

        if prin.drawing_sheet_file_name:
            args = ['kicad-cli', 'pcb', 'export', 'pdf', '--ibt', '--drawing-sheet', prin.drawing_sheet_file_name]
        else:
            args = ['kicad-cli', 'pcb', 'export', 'pdf', '--ibt']

        args += ['-l', layer + ",Edge.Cuts", '--define-var', 'LAYER=' + layer,
                 '--output', output, input_file]

        subprocess.call(args, stdout=subprocess.DEVNULL,
                        stderr=subprocess.STDOUT)

    result = fitz.open()

    for layer in pcb_pdf_layers:

        output_path = os.path.join(os.getcwd(), output_path_pdf)
        output_filename = pcb_name + '_R' + revision + '_' + layer + '.pdf'
        pdf_file = os.path.join(output_path, output_filename)

        with fitz.open(pdf_file) as mfile:
            result.insert_pdf(mfile)

    pdf_save_path = os.path.join(os.getcwd(), output_path_pdf)
    pdf_save_path = os.path.join(pdf_save_path, pcb_name + '_R' + revision +
                                 '_PCB.pdf')

    result.save(pdf_save_path)

    for layer in pcb_pdf_layers:

        if layer == "F.Fab":
            continue

        if layer == "B.Fab":
            continue

        output_path = os.path.join(os.getcwd(), output_path_pdf)
        output_filename = pcb_name + '_R' + revision + '_' + layer + '.pdf'
        pdf_file = os.path.join(output_path, output_filename)
        os.remove(pdf_file)

    print("Done.\n")


def exportErc(input_file, revision):
    print("* Executing electrical rule check...", end="")

    if not os.path.isdir(output_path_rule_checks):
        os.makedirs(output_path_rule_checks)

    args = [
        "kicad-cli",
        "sch",
        "erc",
        "--format",
        "json",
        "--severity-error",
        "--severity-warning",
        "--output",
        os.path.join(
            os.getcwd(),
            output_path_rule_checks,
            project_name + "_R" + revision + "_ERC.json"
        ),
        input_file
    ]

    output = subprocess.check_output(args).decode("utf-8")

    violations = re.search(r"Found (\d+) violations", output)

    num_violations = int(violations.group(1)) if violations else 0

    if not num_violations:
        print("Done.\n")
    else:
        print("Error.")
        print(f"  Found {num_violations} violations.\n")
        raise GeneratorError("Electrical rule check has errors.")


def exportDrc(input_file, revision):
    print("* Executing design rule check...", end="")

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

    output = subprocess.check_output(args).decode("utf-8")

    violations = re.search(r"Found (\d+) violations", output)
    nc_items = re.search(r"Found (\d+) unconnected items", output)
    schpissue = re.search(r"Found (\d+) schematic parity issues", output)

    num_violations = int(violations.group(1)) if violations else 0
    num_nc_items = int(nc_items.group(1)) if nc_items else 0
    num_schpissue = int(schpissue.group(1)) if schpissue else 0

    if not num_violations and not num_nc_items and not num_schpissue:
        print("Done.\n")
    else:
        print("Error.")
        print(f"  Found {num_violations} violations, \
              {num_nc_items} unconnected items and {num_schpissue} schematic \
              parity issues.\n")

        raise GeneratorError("Design rule check has errors.")


def exportGerbers(input_file, copper_layer_list, revision):
    print("* Generating gerber files...", end="")

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

    subprocess.call(args, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)

    for layer in gerber_layers:
        filename = os.path.join(os.getcwd(),
                                output_path_gerber,
                                project_name + "-" + layer.replace(".", "_") +
                                ".gbr")

        new_filename = rreplace(filename, project_name, project_name + "_R" +
                                revision, 1)

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

    subprocess.call(args, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)

    filenames = [
        os.path.join(
            os.getcwd(),
            output_path_gerber,
            project_name + "-NPTH.drl"
        ),
        os.path.join(
            os.getcwd(),
            output_path_gerber,
            project_name + "-NPTH-drl_map.gbr"
        ),
        os.path.join(
            os.getcwd(),
            output_path_gerber,
            project_name + "-PTH.drl"
        ),
        os.path.join(
            os.getcwd(),
            output_path_gerber,
            project_name + "-PTH-drl_map.gbr"
        ),
        os.path.join(
            os.getcwd(),
            output_path_gerber,
            project_name + "-job.gbrjob"
        ),
    ]

    for filename in filenames:

        new_filename = rreplace(
            filename,
            project_name,
            project_name + "_R" + revision, 1
        )

        if os.path.isfile(filename):
            if os.path.isfile(new_filename):
                os.remove(new_filename)

            os.rename(filename, new_filename)

    print("Done.\n")


def exportPickAndPlace(input_file, revision):
    print("* Generating pick and place files...", end="")

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
        os.path.join(
            os.getcwd(),
            output_path_gerber,
            project_name + "_R" + revision + "-top-pos.csv"
        ),
        input_file
    ]

    subprocess.call(args, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)

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
        os.path.join(
            os.getcwd(),
            output_path_gerber,
            project_name + "_R" + revision + "-bottom-pos.csv"
        ),
        input_file
    ]

    subprocess.call(args, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
    print("Done.\n")


def exportStep(input_file, revision):
    print("* Generating 3D step file...", end="")

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
        os.path.join(
            os.getcwd(),
            output_path_3d,
            project_name + "_R" + revision + "_3D.step"
        ),
        input_file
    ]

    subprocess.call(args, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)

    print("Done.\n")


def getFilenameWithouthExtension(filename):
    return os.path.splitext(os.path.basename(filename))[0]


def exportPdfSch(schematic_file_name, revision):
    print("* Generating schematic pdf file...", end="")

    if not os.path.isdir(output_path_pdf):
        os.makedirs(output_path_pdf)

    if prin.drawing_sheet_file_name:
        args = [
            "kicad-cli",
            "sch",
            "export",
            "pdf",
            "--drawing-sheet",
            prin.drawing_sheet_file_name,
            "--output",
            os.path.join(
                os.getcwd(),
                output_path_pdf,
                project_name + "_R" + revision + "_SCH.pdf"
            ),
            schematic_file_name
        ]
    else:
        args = [
            "kicad-cli",
            "sch",
            "export",
            "pdf",
            "--output",
            os.path.join(
                os.getcwd(),
                output_path_pdf,
                project_name + "_R" + revision + "_SCH.pdf"
            ),
            schematic_file_name
        ]

    output = subprocess.call(
        args,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )

    if not output:
        print("Done.\n")
    else:
        print("Error.\n")


def exportBom(schematic_file_name, revision):
    print("* Export bill of materials...", end="")

    if not os.path.isdir(output_path_bom):
        os.makedirs(output_path_bom)

    fields_array = [
        "${QUANTITY}",
        "Reference",
        "Value",
        "Footprint",
        "Description",
        "Tol/Rat/Mat",
        "Manufacturer",
        "Order Number",
        "fit_field",
    ]

    labels_array = [
        "Quantity",
        "Reference",
        "Value",
        "Footprint",
        "Description",
        "Tol/Rat/Mat",
        "Manufacturer",
        "Order Number",
        "fit_field",
    ]

    args = [
        "kicad-cli",
        "sch",
        "export",
        "bom",
        "--ref-range-delimiter",
        "",
        "--fields",
        ",".join(fields_array),
        "--labels",
        ",".join(labels_array),
        "--group-by",
        "Value,Footprint,fit_field",
        "--output",
        os.path.join(
            os.getcwd(),
            output_path_bom,
            project_name + "_R" + revision + "_BOM.csv"
        ),
        schematic_file_name
    ]

    output = subprocess.call(
        args,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )

    if not output:
        print("Done.\n")
    else:
        print("Error.\n")


prin = ProjectInformation()
project_name = getFilenameWithouthExtension(prin.project_file_name)

cli_args = prin.cli_arg_parser.parse_args()

if not cli_args.no_erc or not cli_args.no_drc:
    print("====================== Rule checks =========================\n")

if not cli_args.no_drc:
    try:
        exportDrc(prin.pcb_file_name, prin.revision)
    except GeneratorError as e:
        print(f"Error: {e}")
        print("Terminating...")
        sys.exit()

if not cli_args.no_erc:
    try:
        # put erc into this also
        exportErc(prin.schematic_file_name, prin.revision)
    except GeneratorError as e:
        print(f"Error: {e}")
        print("Terminating...")
        sys.exit()

print("======================= Generate ===========================\n")

# Schematic
exportPdfSch(prin.schematic_file_name, prin.revision)
exportBom(prin.schematic_file_name, prin.revision)

# PCB
exportPdfPcb(prin.pcb_file_name, prin.copper_layers, prin.revision)
exportGerbers(prin.pcb_file_name, prin.copper_layers, prin.revision)
exportPickAndPlace(prin.pcb_file_name, prin.revision)
exportStep(prin.pcb_file_name, prin.revision)
# exportAsm(prin.pcb_file_name, prin.revision)

print("======================= Success ===========================\n")
