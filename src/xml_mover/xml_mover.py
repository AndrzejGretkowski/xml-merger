from argparse import ArgumentParser
from pathlib import Path
from lxml import etree
from warnings import warn
import pandas as pd


def find_files_in_dir(dir: Path):
    xml_files = list(dir.glob('*.xml'))

    if (dir / 'wsad.xlsx').exists():
        excel_files = [Path(dir / 'wsad.xlsx')]
    else:
        excel_files = list(dir.glob('*.xlsx'))

    if len(xml_files) > 1 or len(excel_files) > 1:
        warn(f"There are {len(xml_files)} XML files and {len(excel_files)} Excel files present...")
        input("Press enter to continue...")
        exit(1)

    return xml_files[0], excel_files[0]


def output_name(input_name: Path):
    return input_name.with_stem(f'{input_name.stem}_processed')


def parse_args():
    parser = ArgumentParser(description='This accepts commandline arguments.')

    parser.add_argument(
        '--dir',
        '-d',
        type=Path,
        default=Path('./'),
        help='Working directory.',
    )
    args, _ = parser.parse_known_args()
    return args


if __name__ == '__main__':
    args = parse_args()

    # Paths
    xml_file_path, excel_input_path = find_files_in_dir(args.dir)
    output_file_path = output_name(xml_file_path)

    # Work
    excel = pd.read_excel(excel_input_path, header=None)
    input_data = excel.iloc[:, 0].tolist()

    tree = etree.parse(xml_file_path)
    root = tree.getroot()

    all_codes = list(root.iter("kodPPE"))
    if len(all_codes) != len(input_data):
        warn(f"Wrong number of data: got {len(all_codes)} in XML and {len(input_data)} in Excel")
        input("Press enter to continue...")
        exit(1)

    for code, input_code in zip(all_codes, input_data):
        code.text = input_code

    tree.write(output_file_path, pretty_print=True, xml_declaration=True, encoding=tree.docinfo.encoding, standalone=tree.docinfo.standalone)
