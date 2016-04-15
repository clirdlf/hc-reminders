# Email Sender

## Setup

* Clone the repository
* Install the dependencies (`bundle install`)
* Export the Filemaker Pro database as Excel
* Run the `excel.rb` script
* Profit

If you are working on this script (and not actually wanting to send out emails),
you will need to run `mailcatcher`

```
$ mailcatcher
```

This will set up an SMTP server on your computer that you can access at
[http://127.0.0.1:1080](http://127.0.0.1:1080) (this is also where you quit the
daemon).

## Sending Email

* **Phase 1:** Run `./excel.rb`
* **Phase 2:** ?
* **Phase 3:** Profit

![Phase 2](https://33.media.tumblr.com/0264e5f14b55733b0b6b24aad6a255f9/tumblr_n2qtbhExIg1r4gei2o5_400.gif)

## Editing Templates

Each email has two templates, a `text` and `html` version. The files are
in the `templates` directory and named by the timeframe in which they
are sent. When you need to make changes, **be sure to edit both
versions!!**



