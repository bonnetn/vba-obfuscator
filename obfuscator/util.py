import random
import re
import string


def get_random_string(size: int) -> str:
    return ''.join(random.choice(string.ascii_letters + "_") for _ in range(size))


def get_random_string_of_random_length() -> str:
    var_length = random.randint(10, 14)
    return get_random_string(var_length)


def replace_whole_word(content: str, var_name: str, new_name: str) -> str:
    var_name = re.escape(var_name)
    new_name = re.escape(new_name)
    pattern = r"\b{}\b".format(var_name)
    result = re.sub(pattern, new_name, content)
    return result
