import re


def main(string):
    value = re.sub(r"[\s]", "", string)
    return value


input = "AGOSTO "
print(len(input))
result = main(input)
print(len(result))
