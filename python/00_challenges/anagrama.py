"""
1. Make a function that get two words of string type
2. Validate if those two words are anagrams
3. Return the result depends on the validation
"""

def is_anagram(word_1: str, word_2: str) -> bool:
    """
    Validate is two words are anagrams

    Params:
    a) (string): The first word
    b) (string): The second word

    Returns:
    bool: The result of the validation
    """
    uno = word_1.lower()
    dos = word_2.lower()

    ##If both are equals return False
    if uno == dos:
        return False

    ##Validate if the words are anagrams
    ## 1. Sort the letter into the words
    ## 2. Compare between the words sorted
    uno = "".join(sorted(uno))
    dos = "".join(sorted(dos))

    return uno == dos

uno = "AMOR"
dos = "Roma"
print(is_anagram(uno, dos))
