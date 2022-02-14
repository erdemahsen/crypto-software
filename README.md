## Crypto Auto-Approval Bot
# How to use it?

1- Download Crypto-Bot.py
2- Put your deposit_history.xlsx file into the same folder with Crypto-Bot.py.  
3- Run Crypto-Bot.py  
4- You will get 2 files which are called members.xlsx and notused.xlsx. members.xlsx will have the list of accepted memberships while notused.xlsx will have the txid's of the users who are not accepted but paid some money to the company. (every file is explained in details later)

# What are the files that we have and what do they mean ?:

  There are 3 main files that we will be interested in. These are deposit_history.xlsx, members,xlsx and notused.xlsx.
  
# deposit_history.xlsx : 
  This file consists of 2 sheets that contains data taken from Google forms and Binance. This will be our input. Binance sheet has TXID of the user who sent the money, amount of the money and the type of coin that was sent and more. Google forms sheet has data that was filled by users that who want to get a membership. Important part that they have to fill is txid since that's how we match the members. We take txid that was given at the Google forms sheet and try to find a match at Binance sheet. The proccess will be explained. This file is the input file.

# notused.xlsx :
  This file will be generated when Crypto-Bot is run. It will have a list of txid's that is in Binance form but haven't been used by the program. So that if there is a money that is paid but the person who paid can not be found company take precautions against these cases and find the user who paid the money.
  
# members.xlsx :
  This file is also an output file. If everything goes according to the plan this file will have the info about the users who got approved.

# What does Crypto-Bot.py do?
  As I explained earlier this bot finds the deposit_history.xlsx file and matches datas of it's 2 subsheets. In case of a match it checks out the money type and finds out if the minimum price is paid. Creates 2 new files called members.xlsx and notused.xlsx as output. When program is run there is a temporary database created but when program ends it is deleted. Crypto-Bot.py has it's own date conversion functions so when there is a new member program outputs the ending date of the membership. TRC20USDT, BEP20USDT, BEP20BUSD are the most important variables since they are the minimum limit that should be paid in every coin type other than that there aren't many thing that you should adjust when you use the program.
  I will leave the example xlsx files so that you can try the program with small inputs.

#Ömer Erdem Ahsen
