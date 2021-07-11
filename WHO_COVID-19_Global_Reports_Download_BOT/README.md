## WHO COVID-19 Global Region wise Reports Automation.
---

### Before Note

This is a simple Robotic Process Automation `RPA` project developed to demonstrate the power of ClointFusion Browser Automation functions. This project is presented at `ClointFusion Monthly Hackathon 9.0.` 
> Check out the Explaination of this BOT in [Youtube](https://youtu.be/6Yjvb6nmf24?t=4630)

### What's this project about?

This project is a simple automation of downloading Global COVID-19 Regional wise Reports from [WHO](https://www.who.int/) Website using Python Browser Automation.This project is developed mainly using the `ClointFusion`, a Pythton based RPA package.

For more details on ClointFusion, Please refer [this](https://github.com/clointfusion/clointfusion)


This BOT can do the following
- Download the Regional Wise Global COVID-19 Image Data.
- Segregate the Reports Region wise.
- Zip the reports and mail to the recipient

### PC Requirements

- A Windows/Mac/Linux running PC `Python>=3.8` installed in it.
- Microsoft Outlook Account
- Chrome / Firefox browser.

### How to use?

- Make sure You've installed `Python>=3.8` version
- Install the dependents from `requirements.txt` 
```
pip install -r requirements.txt
```
- Fill your Outlook Credentisls in `credentials.json` at mentioned places.
- Run the python file `main.py`

That's it. Within few moments BOT will do the task.

This BOT can scheduled using Task Scheduler as well.

***NOTE: The Credentils will be stored within your system and this BOT do not store by any means. Triple Check the `credentials.json` before share your modified source code to others.***

### BOT Working Video

https://user-images.githubusercontent.com/47080241/125203090-8a676080-e294-11eb-9237-fbd5a32ea0ea.mp4

--- 

`Author: fharookshaik`\
`Email: fharookshaik.5@gmail.com`\
`GitHub Username: fharookshaik`


<div align='center'>
 
### Happy Coding !!

[![forthebadge](https://forthebadge.com/images/badges/made-with-python.svg)](https://forthebadge.com)

</div>