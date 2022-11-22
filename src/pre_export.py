import xlwings as xw

import functions as fn


def main():
    pre_export()


def pre_export():
    df = fn.get_caller_df().copy()
    mlo_group = df.groupby("MLO")

    vessel = fn.get_caller_df.vessel
    voyage = fn.get_caller_df.voyage
    pol = fn.get_caller_df.pol

    file_name = (
        r"PRE_EXPORT_" + vessel + "_" + str(voyage[:5]) + "_" + pol + "_"
    )  # utelämna slutet för att kompletteras i loop nedan

    fn.save_pre_export_files(mlo_group, file_name, vessel, voyage, pol)


if __name__ == "__main__":
    file_path = fn.get_mock_caller("0109_Bokningsblad.xlsb")
    xw.Book(file_path).set_mock_caller()
    pre_export()
