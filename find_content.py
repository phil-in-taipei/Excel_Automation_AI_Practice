import openpyxl


def get_content_from_excel(filename, start_row=9, end_row=13):
  """
  Extracts content from specified columns and rows of an Excel spreadsheet.

  Args:
    filename: The name of the Excel file.
    start_row: The starting row (inclusive) to extract content from (default: 9).
    end_row: The ending row (exclusive) to extract content from (default: 13).

  Returns:
    A list of lists containing the extracted content.
  """
  try:
    workbook = openpyxl.load_workbook(filename)
  except FileNotFoundError:
    print(f"File not found: {filename}")
    return []

  worksheet = workbook.active
  content = []

  # Extract content from columns C, E, and F
  for row in worksheet.iter_rows(min_row=start_row, max_row=end_row, min_col=3, max_col=6):
    row_content = [cell.value for cell in row]
    content.append(row_content)

  return content


# Example usage
filename = "[TLI]_Blizzard_Learning_program.xlsx"
content = get_content_from_excel(filename)

if content:
  print(content)
else:
  print("No content found in the specified range.")
