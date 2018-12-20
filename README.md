# OutlookTimeSpotPicker
Lets you copy time intervals from your Outlook Calendar to the clipboard. Useful for sharing calendar availability with people outside your organization.

## How to use
After built and installed, right-click on any spot in Outlook's calendar view. It shows up ther as "Copy spot to clipboard". Click on it and it will do just that. This is a sample of how the time spot will be copied:

```
Mon Nov 19th, 02:00am - 02:30am
```
## Modifier keys

If you click on "Copy spot to clipboard" while holding _Shift_, a tab character will be used to separate date and time:
```
Mon Nov 19th__\t__02:00am - 02:30am
```
If you hold down _Control_ it will use a tab character again but will separate date, start time, and end time:
```
Mon Nov 19th__\t__02:00am__\t__02:30am
```

These options are useful when pasting into excel so you can have dates and times in different columns.