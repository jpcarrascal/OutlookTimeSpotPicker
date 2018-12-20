# OutlookTimeSpotPicker
Lets you copy time intervals from your Outlook Calendar to the clipboard. Useful for sharing calendar availability with people outside your organization.

<img src="https://github.com/jpcarrascal/OutlookTimeSpotPicker/blob/master/example.png?raw=true" />

## How to use
After built and installed, right-click on any spot (either an empty one or an appointment) in Outlook's calendar view. It shows up ther as "Copy spot to clipboard". Click on it and it will do just that. This is a sample of how the time spot will be copied:

```
Mon Nov 19th, 02:00am - 02:30am
```
## Modifier keys

If you click on "Copy spot to clipboard" while holding _Shift_, a tab character will be used to separate date and time:
```
Mon Nov 19th\t02:00am - 02:30am
```
If you hold down _Control_ it will use a tab character again but will separate date, start time, and end time:
```
Mon Nov 19th\t02:00am\t02:30am
```

These options are useful, for instance, when pasting into Excel so you can have dates and times in different columns. It actually makes a good companion to my [Excel Auto-Paste](https://github.com/jpcarrascal/ExcelAutoPaste) add-in. 
