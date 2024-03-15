import pandas as pd
import xlrd


def get_num_sheets(path):
    wb = xlrd.open_workbook(path, on_demand=True)
    return len(wb.sheet_names())  
    

def read_df_erwin(path_erwin):
    df_erwin = pd.read_excel(path_erwin)
    return df_erwin


def read_df_superstar(path_superstar):
    df_fct = pd.read_excel(path_superstar)
    return df_fct


def clean_col_name(col_name: str):
    import re

    remove_punctutation = re.sub(r"[^\w\s]", "", col_name)
    clean = remove_punctutation.lower().strip().replace("  ", "_").replace(" ", "_")
    return clean


def clean_df_fct(df_fct):
    return (
        df_fct
        .rename(columns=lambda c: clean_col_name(c))
        .drop(columns=["ts", "cal", "off", "off_to", "link", "username", "date_formula", "no_show", "group", "no_show_reason"])
        # drop repeated header rows
        .loc[lambda df_: df_["email"] != "Email"]
        # drop null email
        .loc[lambda df_: ~df_["email"].isna()]
        # rename date into appt date
        .rename(columns={"date":"appt_date"})
    )


def merge_by_email_then_phone(df_erwin, df_superstar):
    df_merge_email = (df_erwin
        .loc[~df_erwin["email_from"].isna()]
        .merge(
            right=df_superstar,
            left_on="email_from",
            right_on="email",
            validate="many_to_one", 
            how="inner"
        )
    )
    df_merge_phone = (df_erwin
        .loc[~df_erwin["email_from"].isin(df_merge_email["email_from"])]
        .loc[~df_erwin["phone"].isna()]
        .merge(
            right=df_superstar,
            left_on="phone",
            right_on="phone",
            validate="many_to_one", 
            how="inner"
        )
    )
    df_rest = (df_erwin
        .loc[
            ~df_erwin["email_from"].isin(df_merge_email["email_from"]) &\
            ~df_erwin["phone"].isin(df_merge_phone["phone"])
        ]
    )
    df_result = pd.concat([df_merge_email, df_merge_phone, df_rest])
    return df_result