-   [Overview](#overview)
-   [Create the data](#create-the-data)
-   [Write to CSV file:](#write-to-csv-file)
-   [Excel timestamps](#excel-timestamps)

Overview
========

We want to exhibit how to read and write timestaps when interacting with
Excel.

    options(width = 144,
            tidyverse.quiet = TRUE)

    library(tidyverse, quietly = TRUE)
    library(readxl, quietly = TRUE)
    library(purrrlyr, quietly = TRUE)
    library(lubridate, quietly = TRUE)
    library(here, quietly = TRUE)
    library(openxlsx, quietly = TRUE)

Create the data
===============

Create a time sequence we cna wrap our heads around and inspect. We
create two columns, one of which has UTC as its associated timezone, and
the other that casts the same timestamp to America/Chicago:

    # Set up a tibble of timestamps.  Defaults to UTC
    tibble(dt = seq(from=ymd_hms("2018-12-30 00:00:00"), ymd_hms("2019-01-02 00:00:00"), by="6 hours"),
           dt_loc = with_tz(dt, "America/Chicago")) ->
        df

    df %>%
        print(n=Inf)

    ## # A tibble: 13 x 2
    ##    dt                  dt_loc             
    ##    <dttm>              <dttm>             
    ##  1 2018-12-30 00:00:00 2018-12-29 18:00:00
    ##  2 2018-12-30 06:00:00 2018-12-30 00:00:00
    ##  3 2018-12-30 12:00:00 2018-12-30 06:00:00
    ##  4 2018-12-30 18:00:00 2018-12-30 12:00:00
    ##  5 2018-12-31 00:00:00 2018-12-30 18:00:00
    ##  6 2018-12-31 06:00:00 2018-12-31 00:00:00
    ##  7 2018-12-31 12:00:00 2018-12-31 06:00:00
    ##  8 2018-12-31 18:00:00 2018-12-31 12:00:00
    ##  9 2019-01-01 00:00:00 2018-12-31 18:00:00
    ## 10 2019-01-01 06:00:00 2019-01-01 00:00:00
    ## 11 2019-01-01 12:00:00 2019-01-01 06:00:00
    ## 12 2019-01-01 18:00:00 2019-01-01 12:00:00
    ## 13 2019-01-02 00:00:00 2019-01-01 18:00:00

Write to CSV file:
==================

What happens when we write this data to a CSV file?

    write_csv(df, here("dat", "ts.csv"))

Now read it back in *raw* and see what it wrote out. This is what the
text looks like:

    cat(read_file(here("dat", "ts.csv")))

    ## dt,dt_loc
    ## 2018-12-30T00:00:00Z,2018-12-30T00:00:00Z
    ## 2018-12-30T06:00:00Z,2018-12-30T06:00:00Z
    ## 2018-12-30T12:00:00Z,2018-12-30T12:00:00Z
    ## 2018-12-30T18:00:00Z,2018-12-30T18:00:00Z
    ## 2018-12-31T00:00:00Z,2018-12-31T00:00:00Z
    ## 2018-12-31T06:00:00Z,2018-12-31T06:00:00Z
    ## 2018-12-31T12:00:00Z,2018-12-31T12:00:00Z
    ## 2018-12-31T18:00:00Z,2018-12-31T18:00:00Z
    ## 2019-01-01T00:00:00Z,2019-01-01T00:00:00Z
    ## 2019-01-01T06:00:00Z,2019-01-01T06:00:00Z
    ## 2019-01-01T12:00:00Z,2019-01-01T12:00:00Z
    ## 2019-01-01T18:00:00Z,2019-01-01T18:00:00Z
    ## 2019-01-02T00:00:00Z,2019-01-02T00:00:00Z

**Notice that both timestamps were written out in UTC!** Which I really
didn’t expect the first time, to be honest. Which means that when we
read them in with the `read_csv()` function, we lose the property that
we set on `dt_loc` which is that we wanted to see this in
America/Chicago time:

    read_csv(here("dat", "ts.csv")) %>%
        print(n=Inf)

    ## Parsed with column specification:
    ## cols(
    ##   dt = col_datetime(format = ""),
    ##   dt_loc = col_datetime(format = "")
    ## )

    ## # A tibble: 13 x 2
    ##    dt                  dt_loc             
    ##    <dttm>              <dttm>             
    ##  1 2018-12-30 00:00:00 2018-12-30 00:00:00
    ##  2 2018-12-30 06:00:00 2018-12-30 06:00:00
    ##  3 2018-12-30 12:00:00 2018-12-30 12:00:00
    ##  4 2018-12-30 18:00:00 2018-12-30 18:00:00
    ##  5 2018-12-31 00:00:00 2018-12-31 00:00:00
    ##  6 2018-12-31 06:00:00 2018-12-31 06:00:00
    ##  7 2018-12-31 12:00:00 2018-12-31 12:00:00
    ##  8 2018-12-31 18:00:00 2018-12-31 18:00:00
    ##  9 2019-01-01 00:00:00 2019-01-01 00:00:00
    ## 10 2019-01-01 06:00:00 2019-01-01 06:00:00
    ## 11 2019-01-01 12:00:00 2019-01-01 12:00:00
    ## 12 2019-01-01 18:00:00 2019-01-01 18:00:00
    ## 13 2019-01-02 00:00:00 2019-01-02 00:00:00

We can always reset the timezone attribute on a column:

    read_csv(here("dat", "ts.csv")) %>%
        mutate(dt_loc = with_tz(dt_loc, "America/Chicago")) %>%
        print(n=Inf)

    ## Parsed with column specification:
    ## cols(
    ##   dt = col_datetime(format = ""),
    ##   dt_loc = col_datetime(format = "")
    ## )

    ## # A tibble: 13 x 2
    ##    dt                  dt_loc             
    ##    <dttm>              <dttm>             
    ##  1 2018-12-30 00:00:00 2018-12-29 18:00:00
    ##  2 2018-12-30 06:00:00 2018-12-30 00:00:00
    ##  3 2018-12-30 12:00:00 2018-12-30 06:00:00
    ##  4 2018-12-30 18:00:00 2018-12-30 12:00:00
    ##  5 2018-12-31 00:00:00 2018-12-30 18:00:00
    ##  6 2018-12-31 06:00:00 2018-12-31 00:00:00
    ##  7 2018-12-31 12:00:00 2018-12-31 06:00:00
    ##  8 2018-12-31 18:00:00 2018-12-31 12:00:00
    ##  9 2019-01-01 00:00:00 2018-12-31 18:00:00
    ## 10 2019-01-01 06:00:00 2019-01-01 00:00:00
    ## 11 2019-01-01 12:00:00 2019-01-01 06:00:00
    ## 12 2019-01-01 18:00:00 2019-01-01 12:00:00
    ## 13 2019-01-02 00:00:00 2019-01-01 18:00:00

and now we see this in the America/Chicago timezone again.

If we want to control writing out the timestamp in the local format, we
can control that with `strftime()`, which handles converting the
timestamp to a formatted string manually.

    df %>%
        mutate(dt_loc = strftime(dt_loc, "%F %T"))

    ## # A tibble: 13 x 2
    ##    dt                  dt_loc             
    ##    <dttm>              <chr>              
    ##  1 2018-12-30 00:00:00 2018-12-29 18:00:00
    ##  2 2018-12-30 06:00:00 2018-12-30 00:00:00
    ##  3 2018-12-30 12:00:00 2018-12-30 06:00:00
    ##  4 2018-12-30 18:00:00 2018-12-30 12:00:00
    ##  5 2018-12-31 00:00:00 2018-12-30 18:00:00
    ##  6 2018-12-31 06:00:00 2018-12-31 00:00:00
    ##  7 2018-12-31 12:00:00 2018-12-31 06:00:00
    ##  8 2018-12-31 18:00:00 2018-12-31 12:00:00
    ##  9 2019-01-01 00:00:00 2018-12-31 18:00:00
    ## 10 2019-01-01 06:00:00 2019-01-01 00:00:00
    ## 11 2019-01-01 12:00:00 2019-01-01 06:00:00
    ## 12 2019-01-01 18:00:00 2019-01-01 12:00:00
    ## 13 2019-01-02 00:00:00 2019-01-01 18:00:00

with tibbles, it is easy to see the type of each column. Note that `dt`
is of type `<dttm>`, while by using `strftime()`, we’ve converted
`dt_loc` to a character (`<chr>`).

Here we cast `dt_loc` to a datetime format, write it to a file, and read
the raw contents of that file back in. By *read the raw contents in*, I
mean that here I display what the contents of the file looks like if you
opened it in a text editor, before what R does to it when reading into a
data frame and trying to infer that this is a datetime stamp and cast it
to a datetime representation internally in R.

    df %>%
        mutate(dt_loc = strftime(dt_loc, "%F %T")) %>%
        write_csv(here("dat", "ts.csv"))

    cat(read_file(here("dat", "ts.csv")))

    ## dt,dt_loc
    ## 2018-12-30T00:00:00Z,2018-12-29 18:00:00
    ## 2018-12-30T06:00:00Z,2018-12-30 00:00:00
    ## 2018-12-30T12:00:00Z,2018-12-30 06:00:00
    ## 2018-12-30T18:00:00Z,2018-12-30 12:00:00
    ## 2018-12-31T00:00:00Z,2018-12-30 18:00:00
    ## 2018-12-31T06:00:00Z,2018-12-31 00:00:00
    ## 2018-12-31T12:00:00Z,2018-12-31 06:00:00
    ## 2018-12-31T18:00:00Z,2018-12-31 12:00:00
    ## 2019-01-01T00:00:00Z,2018-12-31 18:00:00
    ## 2019-01-01T06:00:00Z,2019-01-01 00:00:00
    ## 2019-01-01T12:00:00Z,2019-01-01 06:00:00
    ## 2019-01-01T18:00:00Z,2019-01-01 12:00:00
    ## 2019-01-02T00:00:00Z,2019-01-01 18:00:00

What happens when we read this in via `read_csv()`, which converts this
into a dataframe:

    read_csv(here("dat", "ts.csv")) %>%
        print(n=Inf)

    ## Parsed with column specification:
    ## cols(
    ##   dt = col_datetime(format = ""),
    ##   dt_loc = col_datetime(format = "")
    ## )

    ## # A tibble: 13 x 2
    ##    dt                  dt_loc             
    ##    <dttm>              <dttm>             
    ##  1 2018-12-30 00:00:00 2018-12-29 18:00:00
    ##  2 2018-12-30 06:00:00 2018-12-30 00:00:00
    ##  3 2018-12-30 12:00:00 2018-12-30 06:00:00
    ##  4 2018-12-30 18:00:00 2018-12-30 12:00:00
    ##  5 2018-12-31 00:00:00 2018-12-30 18:00:00
    ##  6 2018-12-31 06:00:00 2018-12-31 00:00:00
    ##  7 2018-12-31 12:00:00 2018-12-31 06:00:00
    ##  8 2018-12-31 18:00:00 2018-12-31 12:00:00
    ##  9 2019-01-01 00:00:00 2018-12-31 18:00:00
    ## 10 2019-01-01 06:00:00 2019-01-01 00:00:00
    ## 11 2019-01-01 12:00:00 2019-01-01 06:00:00
    ## 12 2019-01-01 18:00:00 2019-01-01 12:00:00
    ## 13 2019-01-02 00:00:00 2019-01-01 18:00:00

This *looks* good, but there is a problem. `read_csv()` assumes for
timestamps that, if there is no indicator for timezone (like the
trailing ‘Z’ in the `dt` column you see above), then it assumes the
timestamp is in UTC.

    read_csv(here("dat", "ts.csv")) %>%
        mutate(dt_loc = with_tz(dt_loc, tz="UTC")) %>%
        mutate(is_equal = dt == dt_loc) %>%
        print(n=Inf)

    ## Parsed with column specification:
    ## cols(
    ##   dt = col_datetime(format = ""),
    ##   dt_loc = col_datetime(format = "")
    ## )

    ## # A tibble: 13 x 3
    ##    dt                  dt_loc              is_equal
    ##    <dttm>              <dttm>              <lgl>   
    ##  1 2018-12-30 00:00:00 2018-12-29 18:00:00 FALSE   
    ##  2 2018-12-30 06:00:00 2018-12-30 00:00:00 FALSE   
    ##  3 2018-12-30 12:00:00 2018-12-30 06:00:00 FALSE   
    ##  4 2018-12-30 18:00:00 2018-12-30 12:00:00 FALSE   
    ##  5 2018-12-31 00:00:00 2018-12-30 18:00:00 FALSE   
    ##  6 2018-12-31 06:00:00 2018-12-31 00:00:00 FALSE   
    ##  7 2018-12-31 12:00:00 2018-12-31 06:00:00 FALSE   
    ##  8 2018-12-31 18:00:00 2018-12-31 12:00:00 FALSE   
    ##  9 2019-01-01 00:00:00 2018-12-31 18:00:00 FALSE   
    ## 10 2019-01-01 06:00:00 2019-01-01 00:00:00 FALSE   
    ## 11 2019-01-01 12:00:00 2019-01-01 06:00:00 FALSE   
    ## 12 2019-01-01 18:00:00 2019-01-01 12:00:00 FALSE   
    ## 13 2019-01-02 00:00:00 2019-01-01 18:00:00 FALSE

If you’ve cast them right the columns should represent the same times,
which you can check by the `==` operator. You can see above that they
don’t.

What’s happened is that you’ve written out the ‘America/Chicago’
formatted timestamp, but there is no timezone info attached to the
columns in the CSV file. By default, `read_csv()` assumes UTC, which
doesn’t work here.

But you can explicitly tell R what timezone a string format should be
interpreted with. We do so below.

And again, we check to see if the timestamps are equivalent, like they
were at the outset, and indeed they are.

    read_csv(here("dat", "ts.csv"),
             col_types = cols(dt = col_datetime(format = ""),
                              dt_loc = col_character())) %>%
        mutate(dt_loc = ymd_hms(dt_loc, tz="America/Chicago")) %>%
        mutate(is_equal = dt == dt_loc) %>%
        print(n=Inf)

    ## # A tibble: 13 x 3
    ##    dt                  dt_loc              is_equal
    ##    <dttm>              <dttm>              <lgl>   
    ##  1 2018-12-30 00:00:00 2018-12-29 18:00:00 TRUE    
    ##  2 2018-12-30 06:00:00 2018-12-30 00:00:00 TRUE    
    ##  3 2018-12-30 12:00:00 2018-12-30 06:00:00 TRUE    
    ##  4 2018-12-30 18:00:00 2018-12-30 12:00:00 TRUE    
    ##  5 2018-12-31 00:00:00 2018-12-30 18:00:00 TRUE    
    ##  6 2018-12-31 06:00:00 2018-12-31 00:00:00 TRUE    
    ##  7 2018-12-31 12:00:00 2018-12-31 06:00:00 TRUE    
    ##  8 2018-12-31 18:00:00 2018-12-31 12:00:00 TRUE    
    ##  9 2019-01-01 00:00:00 2018-12-31 18:00:00 TRUE    
    ## 10 2019-01-01 06:00:00 2019-01-01 00:00:00 TRUE    
    ## 11 2019-01-01 12:00:00 2019-01-01 06:00:00 TRUE    
    ## 12 2019-01-01 18:00:00 2019-01-01 12:00:00 TRUE    
    ## 13 2019-01-02 00:00:00 2019-01-01 18:00:00 TRUE

You can see that although the timestamps look the same, their timezones
are different:

    read_csv(here("dat", "ts.csv"),
             col_types = cols(dt = col_datetime(format = ""),
                              dt_loc = col_character())) %>%
        mutate(dt_loc = ymd_hms(dt_loc, tz="America/Chicago")) %>%
        map(~ attr(.x, "tzone"))

    ## $dt
    ## [1] "UTC"
    ## 
    ## $dt_loc
    ## [1] "America/Chicago"

In short, if you see two timestamp values being equal, but their
timezones are different, they are not the same time.

But you can see how to explicitly control timezone assignment for data
that doesn’t come with timezone info attached to it internally, but you
know what the timezone is.

Excel timestamps
================

R stores data as an epoch timestamp, which is the number of seconds
since Jan 1, 1970. This is a standard across most computer systems.
Excel does things differently. It is supposed to count the number of
days from Jan 1, 1900, and the time part is the equivalent decimal
fraction of a day. So 6pm on a given day has a 0.75 fractional part of
the datetime number for Excel.

I said *supposed* above. That’s because they made a calendar mistake.
They assumed 1900 was a leap year, which it isn’t. The upshot for the
math is that day 1 is not Jan 1, 1900, but effectively, Dec. 31, 1899.
Or more helpfully, day 0 is Dec. 30th, 1899. Just ugly…

Below is an Excel file that has some Excel-native dates, and a string
that is in ISO format. `read_excel()` will parse the native Excel
datetimes to native R datetimes, with UTC as the default timezone. You
have to force the timestamp into a timezone if you want it in another
timezone.

Let’s try writing out excel:

    openxlsx::write.xlsx(df, file=here("dat", "ts-write.xlsx"))

    ## Note: zip::zip() is deprecated, please use zip::zipr() instead

Read it back in:

    read_excel(here("dat", "ts-write.xlsx"))

    ## # A tibble: 13 x 2
    ##    dt                  dt_loc             
    ##    <dttm>              <dttm>             
    ##  1 2018-12-30 00:00:00 2018-12-29 18:00:00
    ##  2 2018-12-30 06:00:00 2018-12-30 00:00:00
    ##  3 2018-12-30 12:00:00 2018-12-30 06:00:00
    ##  4 2018-12-30 18:00:00 2018-12-30 12:00:00
    ##  5 2018-12-31 00:00:00 2018-12-30 18:00:00
    ##  6 2018-12-31 06:00:00 2018-12-31 00:00:00
    ##  7 2018-12-31 12:00:00 2018-12-31 06:00:00
    ##  8 2018-12-31 18:00:00 2018-12-31 12:00:00
    ##  9 2019-01-01 00:00:00 2018-12-31 18:00:00
    ## 10 2019-01-01 06:00:00 2019-01-01 00:00:00
    ## 11 2019-01-01 12:00:00 2019-01-01 06:00:00
    ## 12 2019-01-01 18:00:00 2019-01-01 12:00:00
    ## 13 2019-01-02 00:00:00 2019-01-01 18:00:00

What are the timzones?

    read_excel(here("dat", "ts-write.xlsx")) %>%
        map(~ attr(.x, "tzone"))

    ## $dt
    ## [1] "UTC"
    ## 
    ## $dt_loc
    ## [1] "UTC"

Note that this wrote `dt_loc` out in the timezone you assigned to it
(not default to UTC). But when you read it back in, it assumes that that
local timestamp is actually UTC.

You can correct that:

    read_excel(here("dat", "ts-write.xlsx")) %>%
        mutate(dt_loc = force_tz(dt_loc, tzone="America/Chicago")) %>%
        map(~ attr(.x, "tzone"))

    ## $dt
    ## [1] "UTC"
    ## 
    ## $dt_loc
    ## [1] "America/Chicago"

    # Read in the excel file, force dt_loc to the America/Chicago timezone.
    read_excel(here("dat", "ts-write.xlsx")) %>%
        mutate(dt_loc = force_tz(dt_loc, tzone="America/Chicago")) %>%
        mutate(is_equal = dt == dt_loc) %>%
        print(n=Inf)

    ## # A tibble: 13 x 3
    ##    dt                  dt_loc              is_equal
    ##    <dttm>              <dttm>              <lgl>   
    ##  1 2018-12-30 00:00:00 2018-12-29 18:00:00 TRUE    
    ##  2 2018-12-30 06:00:00 2018-12-30 00:00:00 TRUE    
    ##  3 2018-12-30 12:00:00 2018-12-30 06:00:00 TRUE    
    ##  4 2018-12-30 18:00:00 2018-12-30 12:00:00 TRUE    
    ##  5 2018-12-31 00:00:00 2018-12-30 18:00:00 TRUE    
    ##  6 2018-12-31 06:00:00 2018-12-31 00:00:00 TRUE    
    ##  7 2018-12-31 12:00:00 2018-12-31 06:00:00 TRUE    
    ##  8 2018-12-31 18:00:00 2018-12-31 12:00:00 TRUE    
    ##  9 2019-01-01 00:00:00 2018-12-31 18:00:00 TRUE    
    ## 10 2019-01-01 06:00:00 2019-01-01 00:00:00 TRUE    
    ## 11 2019-01-01 12:00:00 2019-01-01 06:00:00 TRUE    
    ## 12 2019-01-01 18:00:00 2019-01-01 12:00:00 TRUE    
    ## 13 2019-01-02 00:00:00 2019-01-01 18:00:00 TRUE
