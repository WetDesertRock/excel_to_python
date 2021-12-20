import argparse
from excel2python import excel_to_python


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("xlsx_file")
    parser.add_argument("sheet_index", type=int, default=0,
                        help="Which excel sheet to target (starting from 0)")
    parser.add_argument('--alternating-def', "-a", action='append',
                        type=lambda x: tuple(x.split(":")),
                        help="Vertical Range of cell definitions where the header goes first and is then followed by the value. Specified by range syntax: A8:A29")
    parser.add_argument('--horizontal-def', "-z", action='append',
                        type=lambda x: tuple(tuple(v.split(":")) for v in x.split(",")),
                        help="Range of cell definitions defined by header range and value range. Common for tables where calculation is done horizontally and repeated down. Specified by headerRange,valuesRange. For example: B2:Z2,B33:Z33",)

    return parser.parse_args()


def main():
    args = parse_args()
    code = excel_to_python(args.xlsx_file, args.sheet_index, alternating_infos=args.alternating_def,
                           horizontal_infos=args.horizontal_def)
    print(code)


if __name__ == "__main__":
    main()
