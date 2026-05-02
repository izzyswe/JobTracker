# Defining a variable to use as a key
greet = "key_name"

# Putting the dictionary inside the tuple
tuple1 = ({greet: "hello"}, "some other data")

# This will now work because index 0 exists
print(tuple1[0]) # Output: {'key_name': 'hello'}