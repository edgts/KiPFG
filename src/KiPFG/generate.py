#!/usr/bin/env python3

import os
import sexpdata
import subprocess
import argparse
import fitz

parser = argparse.ArgumentParser(
        prog="KiPFG",
        description="Generate production files for KiCad"
)

parser.add_argument('project_name')

args = parser.parse_args()

project_name = args.project_name
output_path_pdf = "PDF"
output_path_gerber = "FAB"
output_path_bom = "BOM"
output_path_3d = "3D"


def sexParse(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        data = f.read()
        parsed_data = sexpdata.loads(data)
        return parsed_data


def getCopperLayerNames(data):

    layers = []

    for element in data:
        if isinstance(element, list):
            if len(element) > 0:
                if isinstance(element[0], sexpdata.Symbol):
                    match str(element[0]):
                        case "layers":
                            for layer in element:
                                if isinstance(layer[2], sexpdata.Symbol):
                                    if layer[1].endswith(".Cu"):
                                        layers.append(layer[1])
    return layers


def exportPdfPcb(input_file, layers, options):

    if not os.path.isdir(output_path_pdf):
        os.makedirs(output_path_pdf)

    pcb_basic_pdf_layers = [
        'F.Paste',
        'B.Paste',
        'F.Silkscreen',
        'B.Silkscreen',
        'F.Mask',
        'B.Mask',
        'Edge.Cuts'
    ]

    pcb_pdf_layers = layers + pcb_basic_pdf_layers

    # generate temporary single pdf files for each layer
    for layer in pcb_pdf_layers:

        output_path = os.path.join(os.getcwd(), output_path_pdf)
        output_filename = project_name + '_' + layer + '.pdf'
        output = os.path.join(output_path, output_filename)

        args = ['kicad-cli', 'pcb', 'export', 'pdf', '--ibt', '-l']
        args += [layer, '--output', output, input_file]

        subprocess.call(args)

    result = fitz.open()

    for layer in pcb_pdf_layers:

        output_path = os.path.join(os.getcwd(), output_path_pdf)
        output_filename = project_name + '_' + layer + '.pdf'
        pdf_file = os.path.join(output_path, output_filename)

        with fitz.open(pdf_file) as mfile:
            result.insert_pdf(mfile)
            os.remove(pdf_file)

    pdf_save_path = os.path.join(os.getcwd(), output_path_pdf)
    pdf_save_path = os.path.join(pdf_save_path, project_name + '_PCB.pdf')
    result.save(pdf_save_path)


def exportDrc(input_file):
    args = [
            "kicad-cli",
            "pcb",
            "drc",
            "--schematic-parity",
            "--format",
            "json",
            "--severity-exclusions",
            input_file
    ]

    subprocess.call(args)


input_file_pcb = project_name + ".kicad_pcb"
input_file_sch = project_name + ".kicad_sch"


parsed_pcb = sexParse(input_file_pcb)

pcb_copper_layer_list = getCopperLayerNames(parsed_pcb)

exportPdfPcb(input_file_pcb, pcb_copper_layer_list, None)
# exportDrc(input_file_pcb)
