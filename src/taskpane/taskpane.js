Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {

    const paragraph = context.document.body.insertParagraph("汉字测试", Word.InsertLocation.end);
    const chinese = context.document.body.insertParagraph("尧舜禹夏商周春秋战国乱悠悠秦汉三国晋统一南朝北朝是对头隋唐五代又十国宋元明清帝王休", Word.InsertLocation.end);
    const paragraph2 = context.document.body.insertParagraph("汉字测试", Word.InsertLocation.end);
    const english = context.document.body.insertParagraph("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", Word.InsertLocation.end);

    paragraph.font.color = "red";
    chinese.font.color = "pink";
    paragraph2.font.color = "red";
    english.font.color = "blue";
    await context.sync();
  });
}
