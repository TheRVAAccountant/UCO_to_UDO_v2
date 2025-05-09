# excel_utils.py

def get_cell_value(cell):
    if cell.data_type == 'f':
        return cell.value
    else:
        return cell.value

def get_calculated_value(cell):
    if cell.data_type == 'f':
        return cell._value
    else:
        return cell.value
