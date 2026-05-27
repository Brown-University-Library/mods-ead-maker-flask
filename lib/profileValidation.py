import os

from dotenv import dotenv_values


IMAGE_ACCESSIBILITY_ALT_TEXT_COLUMN = 'imageAccessibilityAltText'
IMAGE_ACCESSIBILITY_ALT_TEXT_MAXCHARS_SETTING = 'IMAGE_ACCESSIBILITY_ALT_TEXT_MAXCHARS'


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


def getDotenvPath():
    appDirectory = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(os.path.dirname(appDirectory), '.env')


def getSetting(settingName):
    if settingName in os.environ:
        return os.environ.get(settingName)

    settings = dotenv_values(getDotenvPath())
    return settings.get(settingName)


def getPositiveIntegerSetting(settingName):
    value = getSetting(settingName)

    if value is None:
        return None

    try:
        integerValue = int(str(value).strip())
    except ValueError:
        return None

    if integerValue < 1:
        return None

    return integerValue


def getMaxcharsLimit(validation):
    profileMaxchars = validation.get('maxchars')

    if validation.get('col', '') != IMAGE_ACCESSIBILITY_ALT_TEXT_COLUMN:
        return profileMaxchars

    overrideMaxchars = getPositiveIntegerSetting(IMAGE_ACCESSIBILITY_ALT_TEXT_MAXCHARS_SETTING)

    if overrideMaxchars is None:
        return profileMaxchars

    return overrideMaxchars


def getValidationMessage(validation, defaultMessage, context=None):
    message = validation.get('message') or defaultMessage

    if context is None:
        return message

    try:
        return message.format(**context)
    except (KeyError, ValueError):
        return message


def buildError(rowIndex, validation, value, limit=None, defaultMessage='Spreadsheet validation failed.'):
    context = {}

    if limit is not None:
        context['maxchars'] = limit

    error = {
        'row_index': rowIndex,
        'spreadsheet_row': rowIndex + 2,
        'col': validation.get('col', ''),
        'type': validation.get('type', ''),
        'severity': validation.get('severity', 'error'),
        'message': getValidationMessage(validation, defaultMessage, context),
        'value': value,
    }

    if limit is not None:
        error['limit'] = limit

    return error


def validateMaxchars(row, rowIndex, validation):
    column = validation.get('col', '')
    value = normalizeValue(row.get(column, ''))
    maxchars = getMaxcharsLimit(validation)

    if not value or maxchars is None:
        return []

    if len(value) <= maxchars:
        return []

    return [buildError(
        rowIndex,
        validation,
        value,
        maxchars,
        'Value must be {maxchars} characters or fewer.',
    )]


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
