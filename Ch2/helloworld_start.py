#
# Example file for HelloWorld
#

#
# Declaring my first Python function
# This function uses the input function
# to read information typed in by the
# user. It also uses the print command
# which prints information out to the 
# terminal 
#
def main():
    name = input("What is your name? ");
    followUp = "Hello world, introducing... ";
    print(followUp+name);
    print("Nice to meet you, ",name);

#
#
# To limit the execution of functions only to when they are called
# we use the below line to find the function name (__main__) within 
# the file. If it exists then execute the function
#
#
if __name__ == "__main__":
    main();
