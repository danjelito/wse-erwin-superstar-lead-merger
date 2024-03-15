import io

import pandas as pd
import streamlit as st

import module

# title
st.title("Erwin and Superstar FCT Lead Merger")
st.markdown("**A web app used to merge Erwin data with Superstar FCT to get the Appointment Date.**")
st.markdown("*Created by Devan on 2024-03-15*")
st.markdown("*Updated on 2024-03-15*")
st.divider()

# upload file
st.header("Upload Files")

st.markdown("**Note:**")
st.markdown("- Make sure the excel file contains only one sheet.")
st.markdown("- The superstar FCT sheet that is used is 'All Day' sheet.")
st.write("")


erwin_file = st.file_uploader("Upload the Erwin excel file.")
superstar_file = st.file_uploader("Upload the Superstar FCT excel file.")

# if files are uploaded
if erwin_file is not None and superstar_file is not None:

        with st.spinner("In progress..."):
        
            # check if the file contains only one sheet, else throw error
            num_sheet_erwin = module.get_num_sheets(erwin_file)
            if num_sheet_erwin > 1:
                st.error("Erwin excel file contains more than one sheet. Please delete other sheets then reupload.")

            num_sheet_superstar = module.get_num_sheets(superstar_file)
            if num_sheet_superstar > 1:
                st.error("Superstar excel file contains more than one sheet. Please delete other sheets then reupload.")

            if num_sheet_erwin == 1 and num_sheet_superstar == 1:

                # read df
                df_erwin = module.read_df_erwin(erwin_file)
                df_superstar = module.read_df_superstar(superstar_file)

                # clean df
                df_superstar_clean = module.clean_df_superstar(df_superstar)

                # merge df
                df_result = module.merge_by_email_then_phone(df_erwin, df_superstar_clean)

        # save df to one workbook
        st.divider()
        st.header("Download the Result")

        filename = "Result.xlsx"
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            # Write each dataframe to a different worksheet.
            df_result.to_excel(writer, sheet_name="Result", index=False)

            # Close the Pandas Excel writer and output the Excel file to the buffer
            writer.close()

            st.success("Your file is ready to download.", icon="✅")

            # render download button
            st.download_button(
                label="Click to download",
                data=buffer,
                file_name=filename,
                mime="application/vnd.ms-excel",
            )
