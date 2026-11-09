import Cocoa

final class AppDelegate: NSObject, NSApplicationDelegate, NSWindowDelegate {
    var window: NSWindow!
    var child: Process?
    var url: URL?
    var status: NSTextField!
    func applicationDidFinishLaunching(_ notification: Notification) {
        window = NSWindow(contentRect: NSRect(x:0,y:0,width:530,height:235), styleMask:[.titled,.closable,.miniaturizable], backing:.buffered, defer:false)
        window.title = "Mac Mail Rehberi"
        window.center(); window.delegate = self
        let title = NSTextField(labelWithString:"Mac Mail → Firma Rehberi")
        title.font = .boldSystemFont(ofSize:22); title.frame = NSRect(x:24,y:175,width:480,height:30)
        let explanation = NSTextField(wrappingLabelWithString:"Rehber ekranı tarayıcınızda açılır. Mail hesaplarınızı oradan seçin. Bu pencere açık kaldığı sürece yardımcı çalışır. Postalarınız bilgisayarınızda işlenir.")
        explanation.frame = NSRect(x:24,y:99,width:480,height:65)
        status = NSTextField(labelWithString:"Yardımcı başlatılıyor…"); status.frame = NSRect(x:24,y:70,width:480,height:22)
        let open = NSButton(title:"Rehberi Aç",target:self,action:#selector(openDirectory)); open.bezelStyle = .rounded; open.frame=NSRect(x:22,y:22,width:150,height:34)
        let quit = NSButton(title:"Yardımcıyı Kapat",target:self,action:#selector(stop)); quit.bezelStyle = .rounded; quit.frame=NSRect(x:325,y:22,width:180,height:34)
        [title,explanation,status,open,quit].forEach{window.contentView!.addSubview($0)}
        let menu = NSMenu(); let item = NSMenuItem(); menu.addItem(item); let appMenu=NSMenu(); item.submenu=appMenu
        appMenu.addItem(withTitle:"Mac Mail Rehberi’nden Çık",action:#selector(NSApplication.terminate(_:)),keyEquivalent:"q"); NSApp.mainMenu=menu
        window.makeKeyAndOrderFront(nil); NSApp.activate(ignoringOtherApps:true)
        startServer()
    }
    func startServer() {
        let fm=FileManager.default
        let versions=(try? fm.contentsOfDirectory(atPath:"/Library/Frameworks/Python.framework/Versions")) ?? []
        let frameworks=versions.filter{$0.first?.isNumber == true}.sorted{$0.compare($1,options:.numeric) == .orderedDescending}.map{"/Library/Frameworks/Python.framework/Versions/\($0)/bin/python3"}
        let candidates=frameworks+["/opt/homebrew/bin/python3","/usr/local/bin/python3"]
        guard let python=candidates.first(where:{fm.isExecutableFile(atPath:$0)}) else {
            status.stringValue="Python 3 bulunamadı. python.org üzerinden Python 3 kurun."
            return
        }
        let process=Process(); let output=Pipe()
        process.executableURL=URL(fileURLWithPath:python)
        process.arguments=["-I",Bundle.main.resourceURL!.appendingPathComponent("server.py").path]
        process.standardOutput=output; process.standardError=FileHandle.nullDevice
        process.terminationHandler={ [weak self] _ in DispatchQueue.main.async{self?.status.stringValue="Yardımcı durdu. Uygulamayı yeniden açın."}}
        do {try process.run(); child=process} catch {status.stringValue="Yardımcı başlatılamadı. Python kurulumunu kontrol edin.";return}
        DispatchQueue.global().async { [weak self] in
            let data=output.fileHandleForReading.availableData
            guard let value=String(data:data,encoding:.utf8)?.trimmingCharacters(in:.whitespacesAndNewlines),let local=URL(string:value),local.host=="127.0.0.1" else {return}
            DispatchQueue.main.async {self?.url=local; self?.status.stringValue="Hazır · Yalnızca bu bilgisayardan erişilebilir"; self?.openDirectory()}
        }
    }
    @objc func openDirectory(){if let local=url {NSWorkspace.shared.open(local)}}
    @objc func stop(){NSApp.terminate(nil)}
    func windowShouldClose(_ sender:NSWindow)->Bool{stop();return true}
    func applicationWillTerminate(_ notification:Notification){child?.terminate()}
}
let delegate=AppDelegate()
NSApplication.shared.delegate=delegate
NSApplication.shared.run()
