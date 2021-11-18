#
# Example file for working with conditional statements
#


def main():
    x, y = 100, 1000

    # conditional flow uses if, elif, else
    if (x < y):
         st = "x is less than y"
    elif (x == y):
         st = "x is the same as y"
    else:
         st = "x is greater than y"

    #print(st)

    # conditional statements let you use "a if C else b"
    st = "x is than less than y" if (x < y) else "x is greater than or equal to y"

    print(st)

if __name__ == "__main__":
    main()
