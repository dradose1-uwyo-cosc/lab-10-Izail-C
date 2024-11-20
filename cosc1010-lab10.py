# Izail Chamberlain
# UWYO COSC 1010
# Submission Date: 11/19/24
# Lab 10
# Lab Section: 11
# Sources, people worked with, help given to: chatGPT utilized to help fix "Error: The file [insert file] was not found.", and formatting help to improve readability
# N/A

#import modules you will need 

from hashlib import sha256 
from pathlib import Path

def get_hash(to_hash):
    """You can use """
    return sha256(to_hash.encode('utf-8')).hexdigest().upper()



# Files and Exceptions

# For this assignment, you will be writing a program to "crack" a password. You will need to open the file `hash` as this is the password you are trying to "crack."

# To begin, you will need to open the 'rockyou.txt' file:
# - This file contains a list of compromised passwords from the rockyou dump.
# - This is an abridged version, as the full version is quite large.
# - The file contains the plaintext version of the passwords. You will need to hash them to check against the password hash you are trying to crack.
#   - You can use the provided `get_hash()` function to generate the hashes.
#   - Be careful, as "hello" and "hello " would generate a different hash.

# You will need to include a try-except-catch block in your code.
# - The reading of files needs to occur in the try blocks.


# - Read in the value stored within `hash`.
#   - You must use a try and except block.


# Read in the passwords in `rockyou.txt`.
# - Again, you need a try-except-else block.
# Hash each individual password and compare it against the stored hash.
# - When you find the match, print the plaintext version of the password.
# - End your loop.


from hashlib import sha256

def get_hash(to_hash):
    """Returns the SHA-256 hash of a given string in uppercase."""
    return sha256(to_hash.encode('utf-8')).hexdigest().upper()

def crack_password():
    try:
        # Opens/reads hash file
        with open('hash.txt', 'r') as hash_file:
            target_hash = hash_file.read().strip()  # Reads/removes extra spaces/newlines
    except FileNotFoundError:
        print("Error: The file 'hash.txt' was not found.")
        return
    except Exception as e:
        print(f"Error: An unexpected error occurred while reading 'hash.txt': {e}")
        return

    try:
        # Opens/read rockyou.txt file
        with open('rockyou.txt', 'r') as password_file:
            passwords = password_file.readlines()  # Reads lines into a list
    except FileNotFoundError:
        print("Error: The file 'rockyou.txt' was not found.")
        return
    except Exception as e:
        print(f"Error: An unexpected error occurred while reading 'rockyou.txt': {e}")
        return
    else:
        # Cracks the password
        print("Attempting to crack password...")
        for password in passwords:
            password = password.strip()  # Removes extra spaces/newlines
            hashed_password = get_hash(password)
            if hashed_password == target_hash:
                print(f"The password is: '{password}'")
                break
        else:
            print("Password not found in the provided list.")

# Runs program
crack_password()

