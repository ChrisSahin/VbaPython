import sys

if __name__ == "__main__":
    ListArg = sys.argv[1].split(',')
    map_object = map(int, ListArg)
    listInteger = list(map_object)
    sumValue = sum(listInteger)
    print(sumValue)

    sys.exit(sumValue) 