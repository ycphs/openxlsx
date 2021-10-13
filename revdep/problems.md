# structToolbox

<details>

* Version: 1.4.3
* GitHub: NA
* Source code: https://github.com/cran/structToolbox
* Date/Publication: 2021-09-21
* Number of recursive dependencies: 200

Run `revdep_details(, "structToolbox")` for more info

</details>

## Newly broken

*   R CMD check timed out
    

## In both

*   checking installed package size ... NOTE
    ```
      installed size is  6.7Mb
      sub-directories of 1Mb or more:
        R     3.3Mb
        doc   3.1Mb
    ```

# TPP

<details>

* Version: 3.20.1
* GitHub: NA
* Source code: https://github.com/cran/TPP
* Date/Publication: 2021-07-27
* Number of recursive dependencies: 95

Run `revdep_details(, "TPP")` for more info

</details>

## Newly broken

*   checking examples ... ERROR
    ```
    Running examples in ‘TPP-Ex.R’ failed
    The error most likely occurred in:
    
    > ### Name: tppccrCurveFit
    > ### Title: Fit dose response curves
    > ### Aliases: tppccrCurveFit
    > 
    > ### ** Examples
    > 
    > data(hdacCCR_smallExample)
    ...
    
    
     *** caught segfault ***
    address 0x55a800000004, cause 'memory not mapped'
    
    Traceback:
     1: gc(verbose = FALSE)
     2: tppccrCurveFit(data = tppccrTransformed, nCores = 1)
    An irrecoverable exception occurred. R is aborting now ...
    Segmentation fault (core dumped)
    ```

## In both

*   R CMD check timed out
    

*   checking installed package size ... NOTE
    ```
      installed size is 13.5Mb
      sub-directories of 1Mb or more:
        data           1.9Mb
        example_data   8.0Mb
        test_data      1.9Mb
    ```

*   checking dependencies in R code ... NOTE
    ```
    Namespace in Imports field not imported from: ‘broom’
      All declared Imports should be used.
    Unexported objects imported by ':::' calls:
      ‘doParallel:::.options’ ‘mefa:::rep.data.frame’
      See the note in ?`:::` about the use of this operator.
    ```

*   checking R code for possible problems ... NOTE
    ```
    File ‘TPP/R/TPP.R’:
      .onLoad calls:
        packageStartupMessage(msgText, "\n")
    
    See section ‘Good practice’ in '?.onAttach'.
    
    fitSigmoidCCR: no visible global function definition for
      ‘capture.output’
    modelSelector: no visible binding for global variable ‘testHypothesis’
    modelSelector: no visible binding for global variable ‘fitMetric’
    ...
      ‘Protein_ID’
    tpp2dExport: no visible binding for global variable ‘temperature’
    tpp2dImport: no visible binding for global variable ‘temperature’
    tpp2dNormalize: no visible binding for global variable ‘temperature’
    Undefined global functions or variables:
      ..density.. Protein_ID capture.output fitMetric meltcurve_plot
      minMetric temperature testHypothesis
    Consider adding
      importFrom("utils", "capture.output")
    to your NAMESPACE file.
    ```

