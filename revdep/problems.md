# SwathXtend

<details>

* Version: 2.14.0
* GitHub: NA
* Source code: https://github.com/cran/SwathXtend
* Date/Publication: 2021-05-19
* Number of recursive dependencies: 14

Run `revdep_details(, "SwathXtend")` for more info

</details>

## Newly broken

*   R CMD check timed out
    

## In both

*   checking Rd files ... WARNING
    ```
    prepare_Rd: applyttest.Rd:27-28: Dropping empty section \details
    prepare_Rd: applyttest.Rd:36-37: Dropping empty section \note
    prepare_Rd: applyttest.Rd:34-35: Dropping empty section \author
    prepare_Rd: applyttest.Rd:32-33: Dropping empty section \references
    prepare_Rd: applyttestPep.Rd:26-27: Dropping empty section \details
    prepare_Rd: applyttestPep.Rd:35-36: Dropping empty section \note
    prepare_Rd: applyttestPep.Rd:33-34: Dropping empty section \author
    prepare_Rd: applyttestPep.Rd:31-32: Dropping empty section \references
    prepare_Rd: coverage.Rd:7-9: Dropping empty section \description
    prepare_Rd: coverage.Rd:35-37: Dropping empty section \note
    ...
    prepare_Rd: reliabilityCheckSwath.Rd:40-42: Dropping empty section \references
    prepare_Rd: reliabilityCheckSwath.Rd:52-54: Dropping empty section \seealso
    checkRd: (5) reliabilityCheckSwath.Rd:0-64: Must have a \description
    prepare_Rd: swath.means.Rd:7-9: Dropping empty section \description
    prepare_Rd: swath.means.Rd:22-24: Dropping empty section \details
    prepare_Rd: swath.means.Rd:34-36: Dropping empty section \note
    prepare_Rd: swath.means.Rd:31-33: Dropping empty section \author
    prepare_Rd: swath.means.Rd:28-30: Dropping empty section \references
    prepare_Rd: swath.means.Rd:40-42: Dropping empty section \seealso
    checkRd: (5) swath.means.Rd:0-60: Must have a \description
    ```

*   checking installed package size ... NOTE
    ```
      installed size is 352.1Mb
      sub-directories of 1Mb or more:
        files  351.3Mb
    ```

*   checking R code for possible problems ... NOTE
    ```
    reliabilityCheckSwath: warning in venn.diagram(list(seedSwath =
      ds.seed$Peptide, extSwath = ds.ext$Peptide), file = "venn of
      peptide.png", category.names = c("seed", "extended"), fill =
      c("aquamarine1", "chartreuse"), main = paste("Peptides at FDR pass",
      nfdr)): partial argument match of 'file' to 'filename'
    reliabilityCheckSwath: warning in venn.diagram(list(seedSwath =
      ds.seed$Protein, extSwath = ds.ext$Protein), file = "venn of
      protein.png", category.names = c("seed", "extended"), fill =
      c("aquamarine1", "chartreuse"), main = paste("Proteins at FDR pass",
      nfdr)): partial argument match of 'file' to 'filename'
    ...
                 "rainbow", "terrain.colors")
      importFrom("graphics", "abline", "axis", "barplot", "boxplot", "hist",
                 "layout", "legend", "lines", "mtext", "par", "points",
                 "segments", "text")
      importFrom("stats", "aggregate", "as.formula", "cor", "density", "lm",
                 "lowess", "median", "na.omit", "predict", "resid",
                 "residuals", "sd", "t.test")
      importFrom("utils", "data", "read.delim", "read.delim2", "write.csv",
                 "write.table")
    to your NAMESPACE file.
    ```

