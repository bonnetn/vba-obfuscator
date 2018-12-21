import random
import re
import string


def get_random_string(size: int) -> str:
    return ''.join(random.choice(string.ascii_letters) for _ in range(size))


def get_random_string_of_random_length() -> str:
    var_length = random.randint(10, 14)
    return get_random_string(var_length)


def replace_whole_word(content: str, var_name: str, new_name: str) -> str:
    var_name = re.escape(var_name)
    new_name = re.escape(new_name)
    pattern = r"\b{}\b".format(var_name)
    result = re.sub(pattern, new_name, content)
    return result


def get_functions(code):
    """
    Return all functions names defined in the code.
    :param code:
    :return:
    """
    result = re.finditer("(?:Function|Sub)[ ]+(\w+)\(", code, flags=re.M)
    result = map(lambda x: x.group(1), result)
    result = list(result)

    number_of_routines = len(re.findall("(?:End Sub|End Function)", code))
    print(number_of_routines, len(result))
    assert number_of_routines <= len(result), "Could not find the name of all the routines"
    return result


def get_variables_parameters(code):
    """
    Return all parameters (arguments of functions) names defined in the code.
    :param code:
    :return:
    """
    var_names = re.finditer("(?:Function|Sub)[ ]+\w+\(([^\n\)]+)\)", code, flags=re.M)
    var_names = map(lambda x: x.group(1), var_names)
    var_names = map(extract_variables, var_names)

    result = []
    for names in var_names:
        result += names

    return result


def get_variables_const(code):
    """
    Return all variables names defined such as: "Const MyVar..."
    :param code:
    :return:
    """
    var_names = re.finditer("(?:Const|Set)[ ]+(\w+)[ ]+=", code, flags=re.M)
    var_names = map(lambda x: x.group(1), var_names)
    return var_names


def get_variables_defined(code):
    """
    Return all variables names defined such as: "Dim MyVar..."
    :param code:
    :return:
    """
    var_names = re.finditer("^\s*.*(?:Dim|Private|Public)[ ]((?:\w+(?:[ ]+As[ ]+\w+)?[, ]*)+)", code, flags=re.M)
    var_names = map(lambda x: x.group(1), var_names)
    var_names = map(extract_variables, var_names)

    result = []
    for names in var_names:
        result += names

    return result


def extract_variables(text):
    text = text.split(",")
    for i, s in enumerate(text):
        s = s.strip()
        s = s.split()
        if s[0] == "ByVal":
            s = s[1]
        else:
            s = s[0]
        text[i] = s

    return text


def split_var_declaration_from_code(code):
    """ Extract the global variable declarations (prefix) at the beginning of the code from the actual code (suffix)"""
    regex = r'(^(?:\s|.)*?\n)(.*?(?:Sub|Function)(?:\s|.)*$)'
    match = re.match(regex, code)
    assert len(match.groups()) == 2, "Could not find the prefix & suffix of the code"
    return match.group(1), match.group(2)
