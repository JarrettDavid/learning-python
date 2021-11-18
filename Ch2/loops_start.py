#
# Example file for working with loops
#

def main():
  x = 0

  # define a while loop
  #while(x < 5):
  #    print(x)
  #    x+=1

  # define a for loop
  for x in range(0, 5):
       print(x)
       x+=1

  # use a for loop over a collection
  days=["Mon","Tues","Wed","Thur","Fri","Sat","Sun"]
  for d in days:
      print(d)
  
  # use the break and continue statements
  for x in range(5,10):
      if(x==7): break
      print(x)

   #continue will skip over whatever satisfies the condition staement
   #in this case 6 and 8 are skipped while 5, 7 and 9 are printed
  for x in range(5,10):
      if(x % 2 == 0): continue
      print(x)

  #using the enumerate() function to get index. Added condition statements
  # to only print out certain days of the week when condtions are met
  days=["Mon","Tues","Wed","Thur","Fri","Sat","Sun"]
  for i,d in enumerate(days):
      #if(i % 2 == 0): continue
      #print(i,'.',d)
      if(i == 5): break
      print(i,'.',d)
  
if __name__ == "__main__":
  main()
