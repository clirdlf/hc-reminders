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
