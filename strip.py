# first get all lines from file
with open('testfile.txt', 'r') as f:
    lines = f.readlines()

# remove spaces
lines = [line.replace(' ', '') for line in lines]

# finally, write lines in the file
with open('testfile.txt', 'w+') as f:
    f.writelines(lines)

#strip lines
with open("testfile.txt") as f:
    print "".join(line for line in f if not line.isspace())
