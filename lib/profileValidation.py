class ValidationError(Exception):

    def __init__(self, errors):
        super().__init__(formatValidationErrors(errors))
        self.errors = errors


def normalizeValue(value):
    if value is None:
        return ''
    return str(value).strip()


def conditionMatches(row, condition):
    conditionType = condition.get('type', '')

    if conditionType == 'equals':
        column = condition.get('col', '')
        expectedText = normalizeValue(condition.get('text', ''))
        rowText = normalizeValue(row.get(column, ''))
        return rowText.lower() == expectedText.lower()

    return False


def shouldApplyRule(row, validation):
    conditions = validation.get('conditions', [])

    for condition in conditions:
        if not conditionMatches(row, condition):
            return False

    return True


def getValidationMessage(validation, defaultMessage):
    return validation.get('message') or defaultMessage


def buildError(rowIndex, validation, value, limit=None):
    error = {
        'row_index': rowIndex,
        'spreadsheet_row': rowIndex + 2,
        'col': validation.get('col', ''),
        'type': validation.get('type', ''),
        'severity': validation.get('severity', 'error'),
        'message': getValidationMessage(validation, 'Spreadsheet validation failed.'),
        'value': value,
    }

    if limit is not None:
        error['limit'] = limit

    return error


def validateMaxchars(row, rowIndex, validation):
    column = validation.get('col', '')
    value = normalizeValue(row.get(column, ''))
    maxchars = validation.get('maxchars')

    if not value or maxchars is None:
        return []

    if len(value) <= maxchars:
        return []

    return [buildError(rowIndex, validation, value, maxchars)]


def validateRequired(row, rowIndex, validation):
    if not shouldApplyRule(row, validation):
        return []

    column = validation.get('col', '')
    value = normalizeValue(row.get(column, ''))

    if value:
        return []

    return [buildError(rowIndex, validation, value)]


def validateRow(row, rowIndex, validations):
    errors = []

    for validation in validations:
        validationType = validation.get('type', '')

        if validationType == 'maxchars':
            errors.extend(validateMaxchars(row, rowIndex, validation))

        if validationType == 'required':
            errors.extend(validateRequired(row, rowIndex, validation))

    return errors


def validateRows(rows, validations):
    errors = []

    for rowIndex, row in enumerate(rows):
        errors.extend(validateRow(row, rowIndex, validations))

    return errors


def formatValidationErrors(errors):
    lines = []

    for error in errors:
        line = 'Row %s, column "%s": %s' % (
            error.get('spreadsheet_row', ''),
            error.get('col', ''),
            error.get('message', 'Spreadsheet validation failed.'),
        )
        lines.append(line)

    return '\n'.join(lines)
