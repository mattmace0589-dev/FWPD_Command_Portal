const express = require('express');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

// Serve static files (the portal)
app.use(express.static(path.join(__dirname)));

const DATA_DIR = path.join(__dirname, 'data');
const JSON_FILE = path.join(DATA_DIR, 'roster.json');
const CSV_FILE = path.join(__dirname, 'roster.csv');

function ensureDataDir(){
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);
}

function loadJson(){
  ensureDataDir();
  if (!fs.existsSync(JSON_FILE)) {
    // If CSV exists, try to load from it
    if (fs.existsSync(CSV_FILE)) {
      const csv = fs.readFileSync(CSV_FILE, 'utf8');
      const rows = csv.replace(/\r/g,'').split('\n').filter(Boolean);
      const header = rows.shift().split(',');
      const out = rows.map(r=>{
        const cols = r.split(',');
        const obj = {};
        header.forEach((h,i)=> obj[h.trim()]= (cols[i]||'').trim());
        return obj;
      });
      fs.writeFileSync(JSON_FILE, JSON.stringify(out, null, 2));
      return out;
    }
    fs.writeFileSync(JSON_FILE, '[]');
    return [];
  }
  try {
    return JSON.parse(fs.readFileSync(JSON_FILE,'utf8')||'[]');
  } catch(e){
    console.error('Failed to parse roster.json', e);
    return [];
  }
}

function saveJson(data){
  ensureDataDir();
  fs.writeFileSync(JSON_FILE, JSON.stringify(data,null,2));
}

// API: list roster
app.get('/api/roster', (req,res)=>{
  const data = loadJson();
  res.json(data);
});

// API: add roster item
app.post('/api/roster', (req,res)=>{
  const data = loadJson();
  const item = req.body || {};
  // Ensure an ID
  item.ID = item.ID || String(Date.now());
  data.push(item);
  saveJson(data);
  res.status(201).json(item);
});

// API: update by ID
app.put('/api/roster/:id', (req,res)=>{
  const id = req.params.id;
  const data = loadJson();
  const idx = data.findIndex(x=>String(x.ID) === String(id));
  if (idx === -1) return res.status(404).json({error:'Not found'});
  data[idx] = Object.assign({}, data[idx], req.body);
  saveJson(data);
  res.json(data[idx]);
});

// API: delete by ID
app.delete('/api/roster/:id', (req,res)=>{
  const id = req.params.id;
  let data = loadJson();
  const before = data.length;
  data = data.filter(x=>String(x.ID) !== String(id));
  saveJson(data);
  res.json({deleted: before - data.length});
});

app.listen(PORT, ()=>{
  console.log('Server running on http://localhost:'+PORT);
});
