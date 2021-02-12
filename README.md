# What's the frequency kenneth  
### From Time Domain to "Amplitude vs Frequency" using FFT in Excel 
  
## **FFT Engine**

* Part of the code was taken from senuba91 in this [Original Post](https://stackoverflow.com/users/5748328/senuba91)
* The idea came from the Eng, PHD, colleague, friend, Damian Gargicevich.
  
## *ATENTION*  
  
- The Time Data should be in seconds and both columns should be sorted in function of the time. Both ascending and descending sorts are allowed.  
  If you want to use a Time Column with a Date format. A date format is made of days, not seconds. You can find an adaptation for this purpose in PerformAFFTDateFormat.
- The sampling rate should be constant or as constant as possible.  
- This is a matritial formula. This formula should be accepted with "Control + Shift + Enter"
- The number of cells as a result, should be two columns and at least half of the biggest natural number with an integer logarithm with base 2. (if you have 2300 samples, the last 252 samples will be ignored (because is greater than 2048) and the result will have 1024 rows)

## **How can i call the function?**  

=PerformAFFT(TimeInSecsAsRange,DataAsRange)

## **Whats the meaning of the result?**  

First Column: **Frequency in Hertz**  
Second Column: **Amplitud in the same unit as Data**
