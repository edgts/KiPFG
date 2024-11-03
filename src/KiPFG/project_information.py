import argparse
import os
import re
import sys
import sexpdata


class ProjectInformationError(Exception):
    pass


class ProjectInformation:

    def __init__(self) -> None:
        self.current_working_dir = os.getcwd()
        self.project_file_name = None
        self.schematic_file_name = None
        self.pcb_file_name = None
        self.copper_layers = None
        self.schematic_revision = None
        self.pcb_revision = None
        self.revision = None

        self.cli_arg_parser = argparse.ArgumentParser(
            prog="KiPFG",
            description="Generate production files for KiCad"
        )

        self.cli_arg_parser.add_argument(
            '-p',
            '--project-file',
            help="KiCad project file",
            type=str
        )

        self.cli_arg_parser.add_argument(
            '-e',
            '--no-erc',
            help="Disable ERC check",
            action="store_true"
        )

        self.cli_arg_parser.add_argument(
            '-d',
            '--no-drc',
            help="Diesable DRC check",
            action="store_true"
        )

        self.__findProjectFileName()
        self.__readProjectInformation()
        self.printProjectInformation()

    def __terminateWithError(self, e):
        print(f"Error: {e}")
        print("Terminating...")
        sys.exit()

    def __parseSexpressionFromFile(self, file):
        with open(file, 'r', encoding='utf-8') as f:
            data = f.read()
            parsed_data = sexpdata.loads(data)
            return parsed_data

    def __readProjectFileNameFromCli(self, cli_args):

        project_file_name = re.search(r'\b\w+\.kicad_pro\b', cli_args.project_file)

        if project_file_name:
            project_file_name = project_file_name.group().upper()
            [name, extension] = project_file_name.split(".")
            project_file_name = ".".join([name, extension.lower()])

            if project_file_name:
                if os.path.exists(os.path.join(self.current_working_dir, project_file_name)):
                    return project_file_name
                else:
                    raise ProjectInformationError(f"Project file '{project_file_name}' doesn't exist in the current directory.")
        else:
            raise ProjectInformationError(f"Project file '{cli_args.project_file}' seems to have an invalid format. Should be FILENAME.kicad_pro.")

    def __readProjectFileNameAutomatically(self):
        project_file_name = None

        for file in os.listdir(self.current_working_dir):
            if file.endswith(".kicad_pro"):
                project_file_name = file
                break

        if project_file_name:
            return project_file_name
        else:
            raise ProjectInformationError(f"No kicad project file could be found.")

    def __constructSchematicFileName(self):
        if self.project_file_name:
            self.schematic_file_name = self.project_file_name.split(".")[0] + ".kicad_sch"

            if not os.path.exists(os.path.join(os.getcwd(), self.schematic_file_name)):
                raise ProjectInformationError(f"Error: Schematic file '{self.schematic_file_name}' doesn't exist.")

        else:
            raise ProjectInformationError(f"Project file name doesn't exist. Run the function 'findProjectFileName()' first.")

    def __constructPcbFileName(self):
        if self.project_file_name:
            self.pcb_file_name = self.project_file_name.split(".")[0] + ".kicad_pcb"

            if not os.path.exists(os.path.join(os.getcwd(), self.pcb_file_name)):
                raise ProjectInformationError(f"Error: PCB file '{self.pcb_file_name}' doesn't exist.")
        else:
            raise ProjectInformationError(f"Project file name doesn't exist. Run the function 'findProjectFileName()' first.")

    def __getRevisionFromSexp(self, data):
        revision = None

        for element in data:
            if isinstance(element, list):
                if len(element) > 0:
                    if isinstance(element[0], sexpdata.Symbol):
                        if str(element[0]) == "title_block":
                            for title_block_element in element:
                                if isinstance(title_block_element, list):
                                    if len(title_block_element) > 0:
                                        if isinstance(title_block_element[0], sexpdata.Symbol):
                                            if str(title_block_element[0]) == "rev":
                                                revision = title_block_element[1]

        return revision

    def __getCopperLayerInformationFromSexp(self, data):
        layers = []

        for element in data:
            if isinstance(element, list):
                if len(element) > 0:
                    if isinstance(element[0], sexpdata.Symbol):
                        if str(element[0]) == "layers":
                            for layer in element:
                                if isinstance(layer[2], sexpdata.Symbol):
                                    if layer[1].endswith(".Cu"):
                                        layers.append(layer[1])

        return layers

    def __readSchematicInformation(self):
        schematic_sexp_data = self.__parseSexpressionFromFile(self.schematic_file_name)
        revision = self.__getRevisionFromSexp(schematic_sexp_data)

        if revision:
            self.schematic_revision = revision
        else:
            raise ProjectInformationError("Revision information not found in schematic file.")

    def __readPcbInformation(self):
        pcb_sexp_data = self.__parseSexpressionFromFile(self.pcb_file_name)
        revision = self.__getRevisionFromSexp(pcb_sexp_data)
        copper_layers = self.__getCopperLayerInformationFromSexp(pcb_sexp_data)

        if revision:
            self.pcb_revision = revision
        else:
            raise ProjectInformationError("Revision information not found in pcb file.")

        if copper_layers:
            self.copper_layers = copper_layers
        else:
            raise ProjectInformationError("Copper layer information not found in pcb file.")

    def __findProjectFileName(self):

        args = self.cli_arg_parser.parse_args()

        try:
            if args.project_file:
                self.project_file_name = self.__readProjectFileNameFromCli(args)
            else:
                self.project_file_name = self.__readProjectFileNameAutomatically()

            return self.project_file_name

        except ProjectInformationError as e:
            self.__terminateWithError(e)

    def __readProjectInformation(self):
        try:
            self.__constructSchematicFileName()
            self.__constructPcbFileName()

            self.__readSchematicInformation()
            self.__readPcbInformation()

            if self.schematic_revision is self.pcb_revision:
                self.revision = self.schematic_revision
            else:
                raise ProjectInformationError(f"Revision of schematic '{self.schematic_revision}' and pcb '{self.pcb_revision}' don't match.")

        except ProjectInformationError as e:
            self.__terminateWithError(e)

    def printProjectInformation(self):
        print(r"""
======================= KiPFG ==============================
 ___  __        ___      ________    ________  ________
|\  \|\  \     |\  \    |\   __  \  |\  _____\|\   ____\
\ \  \/  /|_   \ \  \   \ \  \|\  \ \ \  \__/ \ \  \___|
 \ \   ___  \   \ \  \   \ \   ____\ \ \   __\ \ \  \  ___
  \ \  \\ \  \   \ \  \   \ \  \___|  \ \  \_|  \ \  \|\  \
   \ \__\\ \__\   \ \__\   \ \__\      \ \__\    \ \_______\
    \|__| \|__|    \|__|    \|__|       \|__|     \|_______|

================= Project Information ======================
        """)
        print(f"Project     : '{self.project_file_name.split(".")[0]}'")
        print(f"Rev         : {self.revision}")
        print(f"Layer Count : {len(self.copper_layers)}")
        print(f"Layer Names : {self.copper_layers[0]}")
        for layer in self.copper_layers[1:]:
            print(f"            : {layer}")

        print()

        print("Files:")
        print(f"PRO  :'{self.project_file_name}'")
        print(f"SCH  :'{self.schematic_file_name}'")
        print(f"PCB  :'{self.pcb_file_name}'\n")

