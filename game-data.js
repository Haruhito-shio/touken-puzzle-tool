const GameData = {
  "characters": [
    {"id":"k1","name":"歌仙兼定"},
    {"id":"k2","name":"一期一振"},
    {"id":"k3","name":"乱藤四郎"},
    {"id":"k4","name":"今剣"},
    {"id":"k5","name":"宗三左文字"},
    {"id":"k6","name":"蜂須賀虎徹"}
  ],
  "materials": [
    {"id":"m1","name":"水"},
    {"id":"m2","name":"木材"},
    {"id":"m3","name":"紙"},
    {"id":"m4","name":"黄色颜料"},
    {"id":"m5","name":"花札"}
  ],
  "breakthroughs": [
    {"fromStage":20,"toStage":30,"requirements":[{"material":"水","count":20}]},
    {"fromStage":30,"toStage":35,"requirements":[{"material":"木材","count":15}]},
    {"fromStage":35,"toStage":40,"requirements":[{"material":"紙","count":8}]},
    {"fromStage":40,"toStage":45,"requirements":[{"material":"黄色颜料","count":12}]}
  ],
  "recipes": [
    {"product":"黄色颜料","ingredient":"水","quantity":2},
    {"product":"黄色颜料","ingredient":"紙","quantity":1}
  ],
  "stages": [
    {"stageId":"3‑5","materials":["水","木材","黄色颜料"]},
    {"stageId":"2‑4","materials":["水","紙"]},
    {"stageId":"4‑1","materials":["木材","花札"]}
  ]
};
