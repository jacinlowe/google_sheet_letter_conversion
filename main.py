""" 
This project is to convert a column name from a spreadsheet to a column number. 
Makes it easier to work with query data in excel and google sheets

"""

charstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

char_dict: dict[str:int] = dict(zip(charstr, range(1, len(charstr) + 1)))

# print(char_dict)


def parse_str(s: str) -> int:
    if len(s) == 1:
        return char_dict[s]

    col_num: int = 0
    for i, c in enumerate(s):
        char_num = char_dict[c]

        if i == len(s) - 1:
            col_num += char_num
        else:
            col_num += char_num * 26
    # col_num -= 1
    return col_num


def get_input():
    s = input("Enter your column name (type '1' to Stop): ")

    return s.upper()


def print_result(string: str, result: int):
    print(f"Column: {string},Column number: {result} \n")


def main():
    while True:
        s = get_input()
        if s == "1":
            print("Exiting...")
            break
        res = parse_str(s)
        print_result(s, res)


if __name__ == "__main__":
    main()
