import { useEffect, useState } from "react";
import "./App.css";

function App() {
  const [retailer, setRetailer] = useState({
    shopName: "",
    ownerName: "",
    email: "",
    phone: "",
    city: "",
  });

  const [purchaseFile, setPurchaseFile] = useState(null);
  const [salesFile, setSalesFile] = useState(null);
  const [result, setResult] = useState(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [mouseLight, setMouseLight] = useState({
  x: 0,
  y: 0,
  visible: false,
});
const [trailDots, setTrailDots] = useState(
  Array.from({ length: 6 }, () => ({
    x: 0,
    y: 0,
    visible: false,
  }))
);
  const purchaseColumns = [
    "ITEM ARTICLE",
    "QTY",
    "SIZE",
    "PURCHASE PRICE",
    "COLOUR",
    "TOTAL AMT",
    "CATEGORY",
    "FIT",
    "MRP",
  ];

  const salesColumns = [
    "ITEM NAME",
    "QTY",
    "MRP",
    "DISCOUNT AMOUNT",
    "COLOUR",
    "NET AMOUNT",
    "CATEGORY",
    "MARGIN",
    "BILL NO",
    "DATE",
    "TAX AMOUNT",
    "SIZE",
  ];

  async function handleAnalyze() {
    if (!purchaseFile || !salesFile) {
      setError("Please upload both Purchase Bill and Sales Report Excel files.");
      return;
    }

    setLoading(true);
    setError("");
    setResult(null);

    const formData = new FormData();
    formData.append("shop_name", retailer.shopName);
    formData.append("owner_name", retailer.ownerName);
    formData.append("email", retailer.email);
    formData.append("phone", retailer.phone);
    formData.append("city", retailer.city);
    formData.append("purchase_file", purchaseFile);
    formData.append("sales_file", salesFile);

    try {
      const response = await fetch("https://your-render-backend-url.onrender.com/analyze", {
        method: "POST",
        body: formData,
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || "Analysis failed");
      }

      setResult(data);
    } catch (err) {
      setError(err.message);
    }

    setLoading(false);
  }
  function handleHeroMouseMove(e) {
  const rect = e.currentTarget.getBoundingClientRect();

  setMouseLight({
    x: e.clientX - rect.left,
    y: e.clientY - rect.top,
    visible: true,
  });
}

function handleHeroMouseLeave() {
  setMouseLight((prev) => ({
    ...prev,
    visible: false,
  }));
}
useEffect(() => {
  let animationFrameId;

  const animateTrail = () => {
    setTrailDots((prevDots) => {
      const newDots = [...prevDots];

      // First dot follows the mouse directly
      newDots[0] = {
        x: mouseLight.x,
        y: mouseLight.y,
        visible: mouseLight.visible,
      };

      // Other dots follow the dot before them
      for (let i = 1; i < newDots.length; i++) {
        newDots[i] = {
          x: newDots[i].x + (newDots[i - 1].x - newDots[i].x) * 0.35,
          y: newDots[i].y + (newDots[i - 1].y - newDots[i].y) * 0.35,
          visible: mouseLight.visible,
        };
      }

      return newDots;
    });

    animationFrameId = requestAnimationFrame(animateTrail);
  };

  animationFrameId = requestAnimationFrame(animateTrail);

  return () => cancelAnimationFrame(animationFrameId);
}, [mouseLight]);

  return (
    <div className="app">
      <section
  className="hero"
  onMouseMove={handleHeroMouseMove}
  onMouseLeave={handleHeroMouseLeave}
>
  <div
    className={`mouseLight ${mouseLight.visible ? "show" : ""}`}
    style={{
      left: `${mouseLight.x}px`,
      top: `${mouseLight.y}px`,
    }}
  ></div>

  {trailDots.map((dot, index) => (
    <div
      key={index}
      className={`trailDot ${dot.visible ? "show" : ""}`}
      style={{
        left: `${dot.x}px`,
        top: `${dot.y}px`,
        opacity: 1 - index * 0.14,
        transform: `translate(-50%, -50%) scale(${1 - index * 0.08})`,
      }}
    ></div>
  ))}

  <div className="bgAnimation"></div>
 
        <div className="bgAnimation">
  <span className="orb orb1"></span>
  <span className="orb orb2"></span>
  <span className="orb orb3"></span>
  <span className="orb orb4"></span>

  <span className="sparkle s1"></span>
  <span className="sparkle s2"></span>
  <span className="sparkle s3"></span>
  <span className="sparkle s4"></span>
  <span className="sparkle s5"></span>
  <span className="sparkle s6"></span>
  <span className="sparkle s7"></span>
  <span className="sparkle s8"></span>
  <span className="sparkle s9"></span>
  <span className="sparkle s10"></span>
  <span className="sparkle s11"></span>
  <span className="sparkle s12"></span>

  <span className="shootingStar star1"></span>
  <span className="shootingStar star2"></span>
  <span className="shootingStar star3"></span>
  <span className="shootingStar star4"></span>

  <span className="cyberLine cLine1"></span>
  <span className="cyberLine cLine2"></span>
  <span className="cyberLine cLine3"></span>

  <span className="scanLine"></span>
  <span className="glowRing ring1"></span>
  <span className="glowRing ring2"></span>
</div>

  <div className="heroContent">

    <div className="titleShowcase">
      <div className="titleGlow"></div>

      <h1 className="projectName">
        Retail <span>Flow</span> AI
      </h1>

      <p className="projectTagline">
        Garment Sales Analysis & Smart Reorder Decision System
      </p>
    </div>

    <div className="benefitsRow">
  <div className="benefitCard">
    <div className="benefitIcon">🧠</div>
    <h4>Smart Reorder</h4>
    <p>Automatically decides which items should be reordered again.</p>
  </div>

  <div className="benefitCard">
    <div className="benefitIcon">📦</div>
    <h4>Stock Tracking</h4>
    <p>Shows sold quantity, remaining stock, and item movement.</p>
  </div>

  <div className="benefitCard">
    <div className="benefitIcon">💰</div>
    <h4>Profit Insights</h4>
    <p>Analyzes margin and profit strength of each garment item.</p>
  </div>

  <div className="benefitCard">
    <div className="benefitIcon">📊</div>
    <h4>Report Dashboard</h4>
    <p>Generates Excel inventory report and interactive dashboard.</p>
  </div>
</div>

    <p className="heroDesc">
      Upload purchase bill and sales report Excel files. Retail Flow AI
      analyzes stock movement, sold quantity, remaining inventory, profit
      margin, and gives smart reorder decisions for retail garment shops.
    </p>

    <div className="heroActions">
      <a href="#analysis" className="primaryBtn">
        Start Analysis
      </a>

      <a
        href="https://your-render-url.onrender.com/dashboard"
        target="_blank"
        rel="noreferrer"
        className="ghostBtn"
      >
      View Dashboard
      </a>
    </div>
  </div>
</section>

      <main className="main" id="analysis">
        <section className="inputSection">
          <div className="formCard">
            <div className="cardHeading">
              <span>01</span>
              <div>
                <h3>Retailer Information</h3>
                <p>Details will be added to the final Excel report.</p>
              </div>
            </div>

            <div className="inputGrid">
              <input
                type="text"
                placeholder="Shop Name"
                value={retailer.shopName}
                onChange={(e) =>
                  setRetailer({ ...retailer, shopName: e.target.value })
                }
              />

              <input
                type="text"
                placeholder="Owner Name"
                value={retailer.ownerName}
                onChange={(e) =>
                  setRetailer({ ...retailer, ownerName: e.target.value })
                }
              />

              <input
                type="email"
                placeholder="Email Address"
                value={retailer.email}
                onChange={(e) =>
                  setRetailer({ ...retailer, email: e.target.value })
                }
              />

              <input
                type="text"
                placeholder="Phone Number"
                value={retailer.phone}
                onChange={(e) =>
                  setRetailer({ ...retailer, phone: e.target.value })
                }
              />

              <input
                type="text"
                placeholder="City / Location"
                value={retailer.city}
                onChange={(e) =>
                  setRetailer({ ...retailer, city: e.target.value })
                }
              />
            </div>
          </div>

          <div className="formCard">
            <div className="cardHeading">
              <span>02</span>
              <div>
                <h3>Upload Excel Files</h3>
                <p>Upload purchase bill and sales report.</p>
              </div>
            </div>

            <div className="uploadGrid">
              <label className="uploadBox">
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={(e) => setPurchaseFile(e.target.files[0])}
                />
                <div className="uploadIcon">📘</div>
                <h4>Purchase Bill</h4>
                <p>{purchaseFile ? purchaseFile.name : "Choose Excel file"}</p>
              </label>

              <label className="uploadBox">
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={(e) => setSalesFile(e.target.files[0])}
                />
                <div className="uploadIcon">📊</div>
                <h4>Sales Report</h4>
                <p>{salesFile ? salesFile.name : "Choose Excel file"}</p>
              </label>
            </div>

            <button className="analyzeBtn" onClick={handleAnalyze}>
              {loading ? "Analyzing Retail Data..." : "Analyze Retail Data"}
            </button>
          </div>
        </section>

        <section className="formatSection">
          <div className="sectionTitle">
            <span>Required Excel Format</span>
            <h3>Cool Column Flow Structure</h3>
            <p>
              Keep your Excel columns in this order so the backend can analyze
              files correctly.
            </p>
          </div>

          <div className="formatFlowBox">
            <div className="flowTitle">Purchase Bill Format</div>
            <div className="flowLine">
              {purchaseColumns.map((col, index) => (
                <div className="flowNode" key={index}>
                  {col}
                </div>
              ))}
            </div>

            <div className="flowTitle salesTitle">Sales Report Format</div>
            <div className="flowLine salesFlow">
              {salesColumns.map((col, index) => (
                <div className="flowNode salesNode" key={index}>
                  {col}
                </div>
              ))}
            </div>
          </div>
        </section>

        {error && <div className="errorBox">{error}</div>}

        {result && (
          <section className="resultSection">
            <div className="resultHeader">
              <div>
                <span>Analysis Completed</span>
                <h3>Retail Flow Report Ready</h3>
                <p>Download Excel report or open interactive dashboard.</p>
              </div>

              <div className="resultActions">
                
                <a
                  href="https://your-render-url.onrender.com/download-report"
                  target="_blank"
                  rel="noreferrer"
                >
                  Download Excel
                </a>
                <a
                  href="https://your-render-url.onrender.com/dashboard"
                  target="_blank"
                  rel="noreferrer"
                >
                  View Dashboard
                </a>
              </div>
            </div>

            <div className="summaryGrid">
              <div>
                <p>Total Items</p>
                <h4>{result.summary.total_items}</h4>
              </div>

              <div>
                <p>High Value</p>
                <h4>{result.summary.high_value_items}</h4>
              </div>

              <div>
                <p>Moderate Value</p>
                <h4>{result.summary.moderate_value_items}</h4>
              </div>

              <div>
                <p>Low Value</p>
                <h4>{result.summary.low_value_items}</h4>
              </div>
            </div>

            <div className="tableBox">
              <table>
                <thead>
                  <tr>
                    <th>Item Name</th>
                    <th>Category</th>
                    <th>Colour</th>
                    <th>Size</th>
                    <th>Purchase Qty</th>
                    <th>Sold Qty</th>
                    <th>Remaining</th>
                    <th>Sales %</th>
                    <th>Value</th>
                    <th>Decision</th>
                  </tr>
                </thead>

                <tbody>
                  {result.preview.map((item, index) => (
                    <tr key={index}>
                      <td>{item["ITEM NAME"]}</td>
                      <td>{item["CATEGORY"]}</td>
                      <td>{item["COLOUR"]}</td>
                      <td>{item["SIZE"]}</td>
                      <td>{item["PURCHASE QTY"]}</td>
                      <td>{item["SOLD QTY"]}</td>
                      <td>{item["REMAINING STOCK"]}</td>
                      <td>{item["SALES PERCENTAGE"]}%</td>
                      <td>{item["VALUE CATEGORY"]}</td>
                      <td>{item["REORDER DECISION"]}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        )}
      </main>
    </div>
  );
}

export default App;