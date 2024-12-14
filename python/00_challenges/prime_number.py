"""
Write a program to know is any number is a prime number or not
"""

def prime_number(number: int) -> bool:
    """
    This function works to kno if the number passed is prime number or not

    Params:
    int: number passed

    Returns:
    bool: confirmation depends on the validation
    """

    ## 1. Knowing if the number is prime number
    ##    1.1 If the number has no any divisor
    ##    1.2 If the number only can be divided by itself
    ## 2. Do the validation using conditionals
    ## 3. Return the result depends on the validation

    if number < 2:
        return False
    """
    If d (any divisor) | n (number) -> n/d es divisor de n 
    Ex: d = 11 n = 99
    d|n -> n/d | n 
    11|99 -> 99/11 | 99
    11|99 -> 9 | 99 = 11
    """
    range_number = int(number**0.5)  ## Square of the number gave
    for i in range(
        2, range_number + 1
    ):  ## +1 make sure that includes the original number
        if number % i == 0:
            return False
    return True


def print_number() -> None:
    ##Print the prime numbers from 1 to 100
    for number in range(1, 101):
        validation = prime_number(number)
        if validation:
            print(number)


print_number()  ##Call the main function
