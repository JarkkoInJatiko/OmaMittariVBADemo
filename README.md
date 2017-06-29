# OmaMittariVBADemo

This Excel file demonstrates how to use OmaMittari gateway in VBA. You can find more information on OmaMittari and how to get energy data in Finland from a REST API in <a href="https://kehitys.omamittari.fi/">OmaMittari Developer Site</a>. This <a href="https://kehitys.omamittari.fi/blog/viesti4">blog post</a> explains the content of this file in more detail. 

OmaMittari is developed by <a href="http://www.jatiko.fi">Jatiko Oy</a>.

## Getting Started

Download OmaMittatiVBADemo.xlsm to your Windows computer. Make sure macros are allowed in Excel. Add your subscription key, username and APIToken in the Helpers module. In order to get them <b>you need to first <a href="https://kehitys.omamittari.fi/signup/">registrate</a> as an OmaMittari developer in order to get your own requiredsubscription.</b>

```
Private Const sUserName As String = "zzzz" '*** Replace this with your UserName
Private Const sAPItoken As String = "zzzz" '*** Replace this with your APItoken
Private Const sSubscription_Key As String = "zzzz" '*** Replace this with your Subscription Key
```

### Prerequisites

This app requires Microsoft Excel (with macros allowed) to be installed in your computer plus credentials (username, APIToken and subscription key)

## Running the tests

Press the buttons in Consumer API or ElectricityNetwork API sheets depending on your username and API Token (whether they are credentials for a customer or for a network company). The macro echoes the API response on the screen.

## Deployment

You can make your own Excel app having a proper UI based on these example requests.

## Contributing

This code is for demonstration purposes only and we don't expect to receive any contributions to it. However feel free to modify it to your own personal or commercial needs.

## Versioning

This file uses OmaMittari API version 1.1.

## Authors

* **Jarkko Lehtonen** - *Initial work* - [JarkkoInJatiko](https://github.com/JarkkoInJatiko/OmaMittariVBADemo)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

Hat tip to those coders whose code was used in this project:
* Barry Dunne in modCrypt module
* Phil Fresle in SHA256 module
* Michael Glaser in JSON module and in cJSONScript module
* Steve McMahon in cStringBuilder
