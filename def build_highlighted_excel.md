def build_highlighted_excel(
    df: pd.DataFrame,
    error_report: pd.DataFrame,
    portal_key: str
) -> bytes:

    output = io.BytesIO()

    # ----------------------------
    # DOLPHIN PORTAL
    # ----------------------------
    if portal_key == "new":

        wb = load_workbook(
            "GroupPolicySample.xlsx"
        )

        worksheet = wb.active

        if worksheet.max_row > 1:
            worksheet.delete_rows(
                2,
                worksheet.max_row
            )

        for row_idx, row in enumerate(
            df.values.tolist(),
            start=2
        ):
            for col_idx, value in enumerate(
                row,
                start=1
            ):
                worksheet.cell(
                    row=row_idx,
                    column=col_idx,
                    value=value
                )

        

        wb.save(output)

        output.seek(0)

        return output.getvalue()

    # ----------------------------
    # GLOBAL PORTAL
    # ----------------------------
    sheet_name = "Global Upload"

    field_map = OLD_PORTAL_FIELD_TO_COLUMN

    yellow_fill = PatternFill(
        fill_type="solid",
        fgColor="FFF2CC"
    )

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name=sheet_name
        )

        worksheet = writer.book[sheet_name]



        column_positions = {
            cell.value: cell.column
            for cell in worksheet[1]
            if cell.value
        }

        if (
            error_report is not None
            and not error_report.empty
        ):

            for _, error in error_report.iterrows():

                export_column = field_map.get(
                    str(error.get("field", ""))
                )

                excel_column = column_positions.get(
                    export_column
                )

                if not excel_column:
                    continue

                try:
                    excel_row = (
                        int(error.get("row_index", 0))
                        + 1
                    )
                except (
                    TypeError,
                    ValueError
                ):
                    continue

                if (
                    excel_row > 1
                    and excel_row <= worksheet.max_row
                ):
                    worksheet.cell(
                        row=excel_row,
                        column=excel_column
                    ).fill = yellow_fill

    

    output.seek(0)

    return output.getvalue()
    print(df.columns.tolist())
