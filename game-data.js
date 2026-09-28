/*
 * 刀剑乱舞材料规划工具 —— 公共游戏数据（单一数据来源）
 * --------------------------------------------------------------
 * 说明（重要，请先读）：
 *  - 因为要“离线双击打开、不花钱、不用服务器”，浏览器不允许 file:// 页面
 *    直接读取 .json 文件，所以这份公共数据放在 game-data.js 里，
 *    页面用 <script> 标签加载。数据结构与 game-data.json 完全一致。
 *  - 这里是唯一的数据来源。代码里没有任何写死的等级、材料、角色。
 *  - 当前内容是【示例数据】，方便先跑通功能。
 *    等有了真实游戏数据，替换本文件里的内容即可，页面会自动跟着变。
 * --------------------------------------------------------------
 * 数据包含 5 块：
 *   characters   可选刀剑角色列表
 *   materials    可选材料列表（base=基础材料，crafted=半成品/合成品）
 *   breakthroughs 突破阶段及材料需求
 *   recipes      合成配方（每一条 = 生产 1 个成品消耗 1 种原料）
 *   stages       关卡及掉落材料清单
 */

window.GAME_DATA = {

  /* ===== 1. 刀剑角色（示例） ===== */
  characters: [
    "歌仙兼定",
    "一期一振",
    "乱藤四郎",
    "今剣",
    "宗三左文字",
    "蜂須賀虎徹",
    "三日月宗近",
    "加州清光"
  ],

  /* ===== 2. 材料（示例）
   *   base    = 基础材料（刷关卡直接得到，不会再被合成展开）
   *   crafted = 半成品/合成品（有合成配方，计算时会递归展开成基础材料）
   * ===== */
  materials: {
    "水":       { type: "base" },
    "木材":     { type: "base" },
    "竹":       { type: "base" },
    "紙":       { type: "base" },
    "石":       { type: "base" },
    "糸":       { type: "base" },
    "黄色颜料": { type: "crafted" },
    "花札":     { type: "crafted" },
    "金箔":     { type: "crafted" }
  },

  /* ===== 3. 突破阶段及材料需求（示例）
   *   fromStage = 起点等级，toStage = 终点等级
   *   needs     = 这一阶段突破需要消耗的材料数量
   * ===== */
  breakthroughs: [
    { fromStage: 20, toStage: 30, needs: { "水": 10, "木材": 5, "紙": 3 } },
    { fromStage: 30, toStage: 35, needs: { "水": 15, "木材": 8, "黄色颜料": 2 } },
    { fromStage: 35, toStage: 40, needs: { "花札": 3, "竹": 10 } },
    { fromStage: 40, toStage: 45, needs: { "金箔": 2, "水": 20 } },
    { fromStage: 45, toStage: 50, needs: { "花札": 5, "黄色颜料": 4 } }
  ],

  /* ===== 4. 合成配方（示例）
   *   每一条表示：生产 1 个 product，需要消耗 ingredient × quantity
   *   一个成品需要多种原料时，就写多条相同 product 的记录。
   * ===== */
  recipes: [
    { product: "黄色颜料", ingredient: "水",   quantity: 1 },
    { product: "黄色颜料", ingredient: "木材", quantity: 2 },
    { product: "花札",     ingredient: "紙",   quantity: 2 },
    { product: "花札",     ingredient: "竹",   quantity: 1 },
    { product: "金箔",     ingredient: "水",   quantity: 1 },
    { product: "金箔",     ingredient: "石",   quantity: 2 }
  ],

  /* ===== 5. 关卡及掉落材料清单（示例）
   *   stage     = 关卡名
   *   materials = 该关卡掉落清单里有的材料
   * ===== */
  stages: [
    { stage: "1-1", materials: ["水", "木材"] },
    { stage: "1-2", materials: ["水", "竹"] },
    { stage: "2-3", materials: ["木材", "紙"] },
    { stage: "3-5", materials: ["水", "木材", "黄色颜料", "花札"] },
    { stage: "4-2", materials: ["竹", "石"] },
    { stage: "5-1", materials: ["紙", "金箔"] },
    { stage: "6-3", materials: ["石", "水", "金箔"] }
  ]

};
//（注：内容由AI生成）
