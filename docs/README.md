
# XLSpeedUp

Excel VBA class that bundles stuff to "speed up" VBA code execution.

This is essentially a republish of Brian Satola's class which can be found at
    <https://chejunkie.com/knowledge-base/speed-up-class-excel-vba/>.
So all credits go to him!

The main reason for this repository is to increase its visibility.

## Features

Bundle some common tricks to speed up VBA code execution. These mainly are

- turn off worksheet calculation
- turn off screen updating
- ignore events

## Prerequisites / Dependencies

None.

## How to install / Getting started

Add `XLSpeedUp.cls` to your project.
Yes, its that simple.

## Usage / Show it in action

A most basic  example is

```vba
Public Sub DoSomething()
    Dim SpeedUp As XLSpeedUp
    Set SpeedUp As New XLSpeedUp
    SpeedUp.TurnOn

    'do something

    SpeedUp.TurnOff
End Sub
```

You can also have a look at the example `XLSpeedUp_demo.xlsm` in the `demo`
folder for a full (dummy) example.

## Running Tests

Yes, [Unit Tests](https://en.wikipedia.org/wiki/Unit_testing) in Excel *are*
possible. For that you need to have the awesome
[Rubberduck](https://rubberduckvba.com/) AddIn installed (and enabled).

1. Open the Visual Basic Editor (<kbd>Alt</kbd>+<kbd>F11</kbd>)
2. Add test modules
   - Code Explorer (of Rubberduck)
     1. Show up the Code Explorer (<kbd>Ctrl</kbd>+<kbd>R</kbd>)
     2. Select the project (or an item of that) to which you want to add the
        test files
     3. Right-click in the Code Explorer and click: Add > Existing files...
     4. Select the file(s) in the `tests` folder and click Open
   - Project Explorer
     1. Show up the Project Explorer (<kbd>Ctrl</kbd>+<kbd>R</kbd>)
        (Hit it twice if the Code Explorer shows up first)
     2. Drag the files in the `tests` folder (in an Explorer window) and drop
        them on the Project in the Project Explorer to which you want to add
        the tests
3. Check that `XLSpeedUp.cls` is present in that project as well. Otherwise
   tests will/should fail.
4. Open Test Explorer (Rubberduck > Unit Tests > Test Explorer)
5. Run the tests by clicking: Run > All Tests

## Used By

This project is used by (at least) these projects:

- <https://github.com/VBA-tools2/SeriesEntriesInCharts>

If you know more, I'll be happy to add them here.

## Known issues and limitations

None that I am aware of.

## Contributing

All contributions are highly welcome!!

If you are new to git/GitHub, please have a look at
    <https://github.com/firstcontributions/first-contributions>
where you will find a lot of useful information for beginners.

I recently was pointed to
    <https://www.conventionalcommits.org>.
which sounds very promising. I'll use them from now on too (and hopefully don't
forget it in a hurry.)

## Further Reading

Here is a collection of links that ...

- Charles Williams Blog: [Making your VBA UDFs Efficient](https://fastexcel.wordpress.com/making-your-vba-udfs-efficient/)
- Microsoft Docs: [Excel performance: Improving calculation performance](https://docs.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-improving-calculation-performance)

## FAQ

1. What are the `'@...` comments good for in the code?
   You should really have a look at the awesome
   [Rubberduck](https://rubberduckvba.com/) project!

## Similar Projects

None that I am aware of.

But if *you* know some, please let me know. Maybe we can combine forces.

## License

[MIT](https://choosealicense.com/licenses/mit/)

<!-- markdownlint-disable-file MD033 -->
