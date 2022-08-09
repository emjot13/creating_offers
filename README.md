This program makes creating offers from excel files easy.

One day my dad came to me and said that he needs to send an offer to one of his clients.</br>
He didn't want to just send them an excel file as it wouldn't look professional.</br>
So I told him that I can write some simple code to make such offers.</br>

The program is supposed to work with specific structure of excel files, that the company my dad works with, is using.</br>
The layout is: name, amount of product left, unit in which item is sold, price.</br>
Such layout is neccessary for program to work without any changes (names of the columns are irrelevant).</br>

Functionalities:

- creating offers:
    User inputs the excel file from which data is pulled.
    Then for each category of products that user wants to add:
        User chooses:
        - The way of looking for products in excel file, by matching a phrase, options:
            - matching in the whole products name
            - matching in the beginning of products name
            - matching for individual words in products name
        - The discount for this category

- modifying offers:
    User inputs the excel file and html file (offer to change) from which data is pulled.
    Then they have following options:
        - deleting products from offer
        - adding products to offer
        - changing the discount of products
    For each of above possibilites the user again can choose any of three ways for looking for products.

- adding the version to the filename if the input filename already exists:
  User creates an offer with the name "batteries"
  Then if they want to create the file with the same name
  There will be added "_X" at the end of the filename, where X is the lowest number such that 
  There isn't a filename with that version number.

The program is written in such way that it's virtually impossible for a user to cause an error. </br>
The user is being asked for an input that can be accounted for as long as necessary.

Below you can see how part of the offer looks like, flashing red background occurs when this product is not available at the moment </br>
And that's also what added caption "CHWILOWO NIEDOSTĘPNE" means in Polish.
</br>
</br>
</br>
<p align="center">
  <img width="460" height="300" src="https://media.giphy.com/media/VQzOTg8Gn778PIiPwu/giphy.gif">
</p>
