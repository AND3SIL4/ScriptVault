"""
Write a function that print the first 50 numbers in the Fibonacci succession
"""

def fibonacci() -> None:
    """
    This function calculate the first 50 numbers in the succession
    and start from 0
    """
    ## 1. Make a list that store the numbers
    ## 2. Initialize the list with the two first numbers
    ## 3. Iterate 49 times
    ## 4. Calculate the next number and append it into the list

    fibonacci: list[int] = [0, 1]

    for i in range(1, 49):
        fibonacci.append(fibonacci[-1] + fibonacci[-2])

    print(fibonacci)

fibonacci() ##Call the main function
