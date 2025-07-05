# changeRDPPort

This software is a maintenance tool for Windows Remote Desktop Server.

After changing the RDP port number in the registry to the argument value, Windows is restarted. Use with caution.

One effective security measure when maintaining a Windows server via Remote Desktop is to specify the IP addresses that are allowed to connect[Firewall - rdp port properties - scope - remote - IP address]. This setting prevents attacks from unspecified sources and eliminates the load on the server.

However, this method causes the problem of being unable to connect if the IP address of the connecting side changes.

This software resolves that issue. If multiple IP addresses are registered in the scope, this software is not necessary, but please consider it as a precautionary measure.

本ソフトは、Windowsリモートデスクトップサーバー保守ツールです。
レジストリのrdpポート番号を引数値に変更後、Windowsを再起動しています。使用は注意してください。

リモートデスクトップでWindowsサーバーを保守する場合の有効なセキュリティ対策として接続を許可するIPアドレスを指定する方法があります[ファイヤーウォール - rdpポートのプロパティ - スコープ - リモートIPアドレス]。この設定は、不特定多数からのアタックを防御し、かつサーバーの負荷を無くします。

しかし、この方法は、接続する側のIPアドレスが変わった場合に接続できなくなる問題を起こします。
このソフトは、その問題を解決します。スコープに複数のIPアドレスを登録している場合は、本ソフトは不要ですが予備対策として検討してください。


# Requirement

Please use it on a Windows machine where PHP is running. PHP is being used to remotely launch this software.
.net framework 4.8 is specified, so please change it as appropriate.

phpが動作しているWindows上で使用してください。遠隔で本ソフトを起動させるためphpを使用しています。
.net framework4.8を指定しているので適宜変更してください。  

# Setup

The following operations are strictly prohibited on operational servers. Remote operations are also prohibited as the connection may be lost.
Please do so at your own risk.

以下の操作は、運用中のサーバーでは厳禁です。また接続が切れるので遠隔での操作も厳禁です。
自己責任でお願いします。

1. Preparation. 準備。

We will decide on two numbers: the default port and the backup port.
The default port for rdp is 3389, but we will change this just to be safe. In what follows, we will **assume** that the default is **50001** and the backup is **50002**.
Check the IP address of the connecting side.

既定ポートと予備ポートの2つの番号を決めます。
rdpの既定ポートは、3389ですが念のためこれも変えます。以下、既定を**50001**、予備を**50002**と**仮定**します。
接続する側のIPアドレスを調べてください。

2.Firewall settings. ファイヤーウォール設定

In the firewall settings, close 3389 and add and open the two ports **50001** and **50002**.

ファイヤーウォールの設定で3389を閉じて**50001**と**50002**の2つを追加し開放します。

![changeRDPPort10](http://teamwind.serveblog.net/github/changeRDPPort/changeRDPPort10.jpg)

Add the IP address to connect to the default scope of **50001**. The configuration for **50002** is not necessary.

既定の**50001**のスコープに接続するIPアドレスを追加してください。**50002**の設定は不要です。

![changeRDPPort1](http://teamwind.serveblog.net/github/changeRDPPort/changeRDPPort1.jpg)

3.Add the static IP masquerade for the router to direct ports **50001** and **50002** to the server.
Please read the manual of your router.

ルーターの静的IPマスカレードに**50001**と**50002**の2つをサーバーに向かうよう追加登録してください。
設定は、お使いのルーターの説明書を読んでください。

4.Change the rdp port in the registry to **50001** and reboot.

レジストリのrdpポートを**50001**に変更して再起動します。

[HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp]PortNumber。

![changeRDPPort2](http://teamwind.serveblog.net/github/changeRDPPort/changeRDPPort2.jpg)

You will now be able to connect to the remote desktop using port **50001**.

以上で新たに**50001**を使用してリモートデスクトップ接続が可能になります。

![changeRDPPort3](http://teamwind.serveblog.net/github/changeRDPPort/changeRDPPort3.jpg)

5.Place this software in any folder. 本ソフトを任意のフォルダに配置します。

If Windows is under UAC management, a dialog will appear when you run it. If you cannot enter it remotely, that's the end of it.
Configure the task scheduler to run this software in order to avoid displaying the UAC confirmation dialog.

WindowsがUAC管理下では、実行時にダイアログが出ます。遠隔で中に入れない場合は、これでジエンドです。
UAC確認ダイアログを出さないためにタスクスケジューラに本ソフトを実行させる設定をします。

6.Add a task in Task Scheduler in Administrative Tools.
If you are in an environment where the UAC confirmation dialog does not appear, this setting is not necessary. Please run the EXE directly from php.

管理ツールのタスクスケジューラでタスクを追加します。UAC確認ダイアログが出ない環境ならばこの設定は必要ありません。直接phpからEXEを実行してください。

![changeRDPPort4](http://teamwind.serveblog.net/github/changeRDPPort/changeRDPPort4.jpg)


![changeRDPPort5](http://teamwind.serveblog.net/github/changeRDPPort/changeRDPPort5.jpg)

Set the path of this software. Use the backup port as an argument. Enter the same path in the start option.

本ソフトのパスを設定します。予備ポートを引数にします。開始オプションに同じパスを入力します。

![changeRDPPort6](http://teamwind.serveblog.net/github/changeRDPPort/changeRDPPort6.jpg)

In the General tab, enter an arbitrary name and check the boxes as shown.

全般タブで任意の名称を入力し図の通りにチェックします。

![changeRDPPort7](http://teamwind.serveblog.net/github/changeRDPPort/changeRDPPort7.jpg)

Enter the administrator password and click OK.

OKボタンで管理者パスワードを入力します。

![changeRDPPort8](http://teamwind.serveblog.net/github/changeRDPPort/changeRDPPort8.jpg)

The task scheduler is now ready.

以上でタスクスケジューラの準備完了です。

7.Create a shortcut for this software and change the link destination in the properties to Task Scheduler.

本ソフトのショートカットを作成しプロパティのリンク先をタスクスケジューラに変更します。

C:\Windows\system32\schtasks.exe /run /tn changerdpport

![changeRDPPort9](http://teamwind.serveblog.net/github/changeRDPPort/changeRDPPort9.jpg)

You can now run this software without the UAC confirmation dialog.
以上でUAC確認ダイアログ無しで本ソフトを実行できます。

8.Prepare PHP to be launched remotely. 遠隔で起動するphpの準備。

I think there are several ways to remotely launch this software, but in my case, I will introduce the method using PHP because it is simple.
The process is simply to launch the shortcut [7] with exec.
This php will be placed in a public location, so please use a complex file name and location.
Also, although it is not included in this sample, please add a password check process.

本ソフトを遠隔で起動する方法は、いくつかあると思いますが私の場合はphpが簡単なので紹介します。
処理は、execで[7]のショートカットを起動するだけです。
このphpは、公開する場所に置くのでファイル名と配置する場所は、複雑な名称にしてください。
また、サンプルにはありませんがパスワード入力処理を追加してください。

	<?PHP
		//Please change the path to match your environment. パスは設置環境に合わせてください
		exec('C:\anydir\changeRDPPort - Shortcut.lnk', $out, $ret);

	?>

# How It Works

If the connecting IP address changes, the firewall scope will prevent access to the default port.

接続する側のIPが変わった場合にファイヤーウォールのスコープによって既定ポートに入れなくなります。

1.php start php呼び出し

[8] Open the php in a browser. This software will start and change the port to the spare port, then restart the server.
After restarting, the spare port that is not set in the scope will open.

[8]で設置したphpをブラウザで開きます。本ソフトが起動して予備ポートに変更後サーバーが再起動します。
再起動後スコープ設定していない予備ポートが開きます。

2.Connect remote desktop リモートデスクトップで接続する

Connect via a spare port.

予備ポートで接続します。

![changeRDPPort11](http://teamwind.serveblog.net/github/changeRDPPort/changeRDPPort11.jpg)

3.Scope the firewall to the new IP.

ファイヤーウォールのスコープに新しいIPを設定します。

![changeRDPPort1](http://teamwind.serveblog.net/github/changeRDPPort/changeRDPPort1.jpg)

4.Change the rdp port in the registry back to **50001** and reboot.

レジストリのrdpポートを**50001**に戻して再起動します。


That is all. The act of remotely rebooting the server carries risks. Please implement this content after establishing sufficient testing and operational backups.

以上です。遠隔でサーバーを再起動する行為は、リスクを伴います。十分な試験と運用のバックアップを確立したうえで本内容を導入してください。

# License

MIT license. Copyright: Teamwind.

# Note

There may be bugs. Use at your own risk. Also, modify the code accordingly.
If you have any requests, please email us. 

バグがあるかもしれません。自己責任でご利用ください。また適宜コード変更してください。
ご要望等がございましたらメール下さい。


(Google Translate)

