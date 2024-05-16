import pandas as pd
import sys


# Constants
HTML_STRING_HEAD = '<table role="presentation" border="1" width="100%" cellspacing="0">\n'
HTML_STRING_TAIL = '</table>\n'
PLACEHOLDER_FALLBACK = "Link"

# CSS Styles
CELL_STYLE = "text-align: center; padding: 12px 8px;"
LINK_BTN_STYLE = '''
    display: inline-block;
    outline: none;
    cursor: pointer;
    font-weight: 500;
    border: 1px solid transparent;
    border-radius: 8px;
    height: 36px;
    line-height: 34px;
    font-size: 14px;
    color: #ffffff;
    background-color: #007c89;
    padding: 0 18px;
'''

LINK_TAG_STYLE = "text-decoration: none; color: white;"


def convert_excel_to_email_html(input_path: str, output_path: str):
    dataframe1 = pd.read_excel(input_path, header=None)

    # Extract header row
    header_row = dataframe1.iloc[0]
    number_of_columns = len(header_row)

    TEXT_COLUMN_WIDTH = 100 // number_of_columns

    html_string = ""

    # Extract link placeholders
    placeholder_row = dataframe1.iloc[1]

    html_string += HTML_STRING_HEAD

    # Create html table header
    html_string += '<tr>\n'
    for index, val in header_row.items():
        # Set fixed width if text column
        if (pd.isna(placeholder_row[index])):
            html_string += f'<th style="width:{
                TEXT_COLUMN_WIDTH}%; {CELL_STYLE}">{val}</th>\n'
        else:
            html_string += f'<th style="{CELL_STYLE}">{val}</th>\n'

    html_string += '</tr>\n'

    # Add the rest cell values
    for index, row in dataframe1.iloc[2:].iterrows():
        html_string += f'<tr>\n'
        for col_index, val in row.items():
            cell_value = ""
            placeholder = placeholder_row[col_index]

            if (not pd.isna(val)):
                if (isinstance(val, str)):
                    if (val.startswith('https://') or val.startswith('http://')):
                        cell_value = val

                        html_string += f'<td style="{CELL_STYLE}">\n'
                        html_string += f'<button style="{LINK_BTN_STYLE}">\n'
                        html_string += f'<a style="{
                            LINK_TAG_STYLE}" href="{cell_value}">'

                        if ((not pd.isna(placeholder)) and isinstance(placeholder, str)):
                            html_string += f'{placeholder}'
                        else:
                            html_string += PLACEHOLDER_FALLBACK

                        html_string += '</a>\n'
                        html_string += '</button>\n'
                        html_string += '</td>\n'
                    else:
                        cell_value = val

                        html_string += f'<td style="{CELL_STYLE}">\n'
                        html_string += f'{cell_value}\n'
                        html_string += '</td>\n'
                else:
                    html_string += f'<td style="{CELL_STYLE}">\n'
                    html_string += "</td>\n"
            else:
                html_string += f'<td style="{CELL_STYLE}">\n'
                html_string += '</td>\n'

        html_string += '</tr>\n'

    html_string += HTML_STRING_TAIL
    with open(output_path + ".txt", 'w') as text_file:
        text_file.write(html_string)


def main():
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print("Error: Incorrect number of arguments. Please provide at least one argument and at most two arguments.")
        sys.exit(1)  # Exit with a non-zero status code to indicate an error

    # Extract the first argument
    argument1 = sys.argv[1]

    # If the second argument is provided, extract it
    argument2 = sys.argv[2] if len(sys.argv) == 3 else 'output'

    convert_excel_to_email_html(argument1, argument2)


main()
