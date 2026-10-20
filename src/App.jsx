import defaultPosts from './posts_data.json';
import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Search, Bookmark, ExternalLink, Moon, Sun, 
  Upload, Download, Plus, BookOpen, Trash2, 
  Edit3, X
} from 'lucide-react';

// IndexedDB Storage Helpers
const DB_NAME = 'MyKnowledgeHubDB';
const STORE_NAME = 'posts';
const DB_VERSION = 1;

function openDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: 'id' });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function getAllPostsFromDB() {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const store = tx.objectStore(STORE_NAME);
    const req = store.getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

async function savePostsToDB(posts) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    posts.forEach((item) => store.put(item));
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function deletePostFromDB(id) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    store.delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

const CATEGORIES = [
  'All',
  'Oracle HCM',
  'Core HR',
  'BIP',
  'OTBI',
  'Fast Formulas',
  'HDL',
  'OIC',
  'SQL',
  'REST APIs',
  'AI Agents',
  'Cloud Technology',
  'Career',
  'Other'
];

function detectCategory(text = '') {
  const t = text.toLowerCase();
  if (t.includes('fast formula') || t.includes('fast formulas') || t.includes('formula')) return 'Fast Formulas';
  if (t.includes('bip') || t.includes('bi publisher') || t.includes('bipublisher')) return 'BIP';
  if (t.includes('otbi') || t.includes('transactional business intelligence')) return 'OTBI';
  if (t.includes('hdl') || t.includes('hcm data loader') || t.includes('data loader')) return 'HDL';
  if (t.includes('oic') || t.includes('oracle integration cloud') || t.includes('integration cloud')) return 'OIC';
  if (t.includes('rest api') || t.includes('rest apis') || t.includes('bpm rest') || t.includes('api ')) return 'REST APIs';
  if (t.includes('core hr') || t.includes('worker') || t.includes('payroll') || t.includes('assignment')) return 'Core HR';
  if (t.includes('ai agent') || t.includes('genai') || t.includes('chatgpt') || t.includes('llm') || t.includes('generative ai')) return 'AI Agents';
  if (t.includes('sql') || t.includes('query') || t.includes('queries')) return 'SQL';
  if (t.includes('oracle hcm') || t.includes('fusion hcm') || t.includes('hcm')) return 'Oracle HCM';
  if (t.includes('cloud') || t.includes('oci') || t.includes('fusion') || t.includes('oracle')) return 'Cloud Technology';
  if (t.includes('career') || t.includes('interview') || t.includes('resume') || t.includes('hiring')) return 'Career';
  return 'Other';
}

export default function App() {
  const [posts, setPosts] = useState([]);
  const [selectedCategory, setSelectedCategory] = useState('All');
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const [sortBy, setSortBy] = useState('newest');
  const [darkMode, setDarkMode] = useState(false);
  const [activeModal, setActiveModal] = useState(null);
  const [editingPost, setEditingPost] = useState(null);
  const [selectedPostDetail, setSelectedPostDetail] = useState(null);
  const [importStats, setImportStats] = useState(null);
  const fileInputRef = useRef(null);

  // With this:
useEffect(() => {
  getAllPostsFromDB().then((data) => {
    if (data && data.length > 0) {
      setPosts(data);
    } else {
      setPosts(defaultPosts);
      savePostsToDB(defaultPosts);
    }
  });
}, []);

  useEffect(() => {
    if (darkMode) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }, [darkMode]);

  const handleUpdatePost = async (updatedPost) => {
    const nextPosts = posts.map((p) => (p.id === updatedPost.id ? updatedPost : p));
    setPosts(nextPosts);
    await savePostsToDB([updatedPost]);
    if (selectedPostDetail && selectedPostDetail.id === updatedPost.id) {
      setSelectedPostDetail(updatedPost);
    }
  };

  const toggleFavorite = (post, e) => {
    e?.stopPropagation();
    handleUpdatePost({ ...post, favorite: !post.favorite });
  };

  const cycleStatus = (post, e) => {
    e?.stopPropagation();
    const nextStatus = post.readStatus === 'unread' ? 'reading' : post.readStatus === 'reading' ? 'completed' : 'unread';
    handleUpdatePost({ ...post, readStatus: nextStatus });
  };

  const handleDeletePost = async (id, e) => {
    e?.stopPropagation();
    if (!confirm('Are you sure you want to delete this post?')) return;
    setPosts((prev) => prev.filter((p) => p.id !== id));
    await deletePostFromDB(id);
    if (selectedPostDetail?.id === id) setSelectedPostDetail(null);
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    const ext = file.name.split('.').pop().toLowerCase();

    reader.onload = async (evt) => {
      try {
        let importedRecords = [];
        if (ext === 'json') {
          const parsed = JSON.parse(evt.target.result);
          importedRecords = Array.isArray(parsed) ? parsed : [parsed];
        } else {
          const data = new Uint8Array(evt.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          importedRecords = XLSX.utils.sheet_to_json(firstSheet);
        }

        let addedCount = 0;
        let updatedCount = 0;
        const currentPostsMap = new Map(posts.map((p) => [p.postLink || p.id, p]));

        const processed = importedRecords.map((item, idx) => {
          const postLink = item['Post Link'] || item.postLink || item.url || item.link || '';
          const author = item['Poster Name'] || item.author || item.poster || 'Unknown';
          const text = item['Text'] || item.text || item.content || item.summary || '';
          const existing = postLink ? currentPostsMap.get(postLink) : null;

          const tags = (text.match(/#\w+/g) || []).map((t) => t.substring(1)).slice(0, 5);
          const extractedLinks = [item['Extracted Link 1'], item['Extracted Link 2'], item['Extracted Link 3']].filter(Boolean);

          if (existing) {
            updatedCount++;
            return {
              ...existing,
              author: author || existing.author,
              text: text || existing.text,
              category: existing.category || detectCategory(text),
              extractedLinks: extractedLinks.length ? extractedLinks : existing.extractedLinks
            };
          } else {
            addedCount++;
            return {
              id: `post-${Date.now()}-${idx}`,
              postId: item['Post ID'] || `${Date.now()}-${idx}`,
              postLink,
              author,
              title: text.split('\n')[0].substring(0, 100) || `Post #${idx + 1}`,
              text,
              category: item.category || detectCategory(text),
              tags,
              extractedLinks,
              favorite: false,
              readStatus: 'unread',
              notes: '',
              summary: '',
              takeaways: [],
              importedAt: new Date().toISOString()
            };
          }
        });

        const mergedMap = new Map();
        posts.forEach((p) => mergedMap.set(p.postLink || p.id, p));
        processed.forEach((p) => mergedMap.set(p.postLink || p.id, p));
        const mergedArray = Array.from(mergedMap.values());

        setPosts(mergedArray);
        await savePostsToDB(mergedArray);
        setImportStats({ total: importedRecords.length, added: addedCount, updated: updatedCount });
      } catch (err) {
        alert('Error parsing file: ' + err.message);
      }
    };

    if (ext === 'json') {
      reader.readAsText(file);
    } else {
      reader.readAsArrayBuffer(file);
    }
  };

  const handleExportBackup = () => {
    const dataStr = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(posts, null, 2));
    const dlAnchor = document.createElement('a');
    dlAnchor.setAttribute('href', dataStr);
    dlAnchor.setAttribute('download', `my_knowledge_hub_backup_${new Date().toISOString().slice(0, 10)}.json`);
    dlAnchor.click();
  };

  const filteredPosts = useMemo(() => {
    return posts
      .filter((post) => {
        const matchesCategory = selectedCategory === 'All' || post.category === selectedCategory;
        const matchesStatus =
          statusFilter === 'all'
            ? true
            : statusFilter === 'favorite'
            ? post.favorite
            : post.readStatus === statusFilter;

        const q = searchQuery.toLowerCase();
        const matchesSearch =
          !searchQuery ||
          post.author.toLowerCase().includes(q) ||
          post.text.toLowerCase().includes(q) ||
          post.notes?.toLowerCase().includes(q) ||
          (post.tags && post.tags.some((tag) => tag.toLowerCase().includes(q)));

        return matchesCategory && matchesStatus && matchesSearch;
      })
      .sort((a, b) => {
        if (sortBy === 'newest') return (b.importedAt || '').localeCompare(a.importedAt || '');
        if (sortBy === 'oldest') return (a.importedAt || '').localeCompare(b.importedAt || '');
        if (sortBy === 'author') return a.author.localeCompare(b.author);
        return 0;
      });
  }, [posts, selectedCategory, statusFilter, searchQuery, sortBy]);

  return (
    <div className="flex h-screen bg-slate-50 dark:bg-slate-950 font-sans text-slate-800 dark:text-slate-100 overflow-hidden">
      {/* Sidebar */}
      <aside className="w-64 border-r border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 flex flex-col shrink-0">
        <div className="p-4 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between">
          <div className="flex items-center space-x-2">
            <div className="w-8 h-8 rounded-lg bg-blue-600 flex items-center justify-center text-white font-bold text-lg">
              Ω
            </div>
            <span className="font-bold text-lg text-slate-900 dark:text-white">Knowledge Hub</span>
          </div>
          <button
            onClick={() => setDarkMode(!darkMode)}
            className="p-1.5 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-500"
          >
            {darkMode ? <Sun size={18} /> : <Moon size={18} />}
          </button>
        </div>

        {/* Quick Filters */}
        <div className="p-3 border-b border-slate-200 dark:border-slate-800 space-y-1 text-sm font-medium">
          <button
            onClick={() => setStatusFilter('all')}
            className={`w-full flex items-center justify-between px-3 py-2 rounded-lg ${
              statusFilter === 'all' ? 'bg-blue-50 text-blue-600 dark:bg-blue-900/40 dark:text-blue-300' : 'hover:bg-slate-100 dark:hover:bg-slate-800'
            }`}
          >
            <span>All Posts</span>
            <span className="text-xs font-semibold px-2 py-0.5 rounded-full bg-slate-200 dark:bg-slate-700">{posts.length}</span>
          </button>
          <button
            onClick={() => setStatusFilter('favorite')}
            className={`w-full flex items-center justify-between px-3 py-2 rounded-lg ${
              statusFilter === 'favorite' ? 'bg-amber-50 text-amber-600 dark:bg-amber-900/40 dark:text-amber-300' : 'hover:bg-slate-100 dark:hover:bg-slate-800'
            }`}
          >
            <span className="flex items-center gap-2"><Bookmark size={16} /> Favorites</span>
            <span className="text-xs font-semibold px-2 py-0.5 rounded-full bg-slate-200 dark:bg-slate-700">
              {posts.filter((p) => p.favorite).length}
            </span>
          </button>
        </div>

        {/* Categories list */}
        <div className="flex-1 overflow-y-auto p-3 space-y-1">
          <p className="text-xs font-semibold uppercase text-slate-400 px-3 mb-2 tracking-wider">Categories</p>
          {CATEGORIES.map((cat) => {
            const count = cat === 'All' ? posts.length : posts.filter((p) => p.category === cat).length;
            const active = selectedCategory === cat;
            return (
              <button
                key={cat}
                onClick={() => setSelectedCategory(cat)}
                className={`w-full flex items-center justify-between px-3 py-1.5 rounded-md text-sm font-medium transition ${
                  active
                    ? 'bg-blue-600 text-white font-semibold'
                    : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'
                }`}
              >
                <span className="truncate">{cat}</span>
                <span className={`text-xs px-1.5 py-0.5 rounded-full ${active ? 'bg-blue-700 text-white' : 'bg-slate-100 dark:bg-slate-800 text-slate-500'}`}>
                  {count}
                </span>
              </button>
            );
          })}
        </div>

        {/* Action buttons */}
        <div className="p-3 border-t border-slate-200 dark:border-slate-800 space-y-2">
          <button
            onClick={() => setActiveModal('import')}
            className="w-full flex items-center justify-center gap-2 px-3 py-2 rounded-lg text-sm font-medium border border-slate-300 dark:border-slate-700 hover:bg-slate-100 dark:hover:bg-slate-800"
          >
            <Upload size={16} /> Import Posts
          </button>
          <button
            onClick={handleExportBackup}
            className="w-full flex items-center justify-center gap-2 px-3 py-2 rounded-lg text-sm font-medium text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800"
          >
            <Download size={16} /> Backup JSON
          </button>
        </div>
      </aside>

      {/* Main Container */}
      <main className="flex-1 flex flex-col min-w-0 overflow-hidden">
        {/* Top Navbar */}
        <header className="h-16 border-b border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 px-6 flex items-center justify-between gap-4 shrink-0">
          <div className="flex-1 relative max-w-xl">
            <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
            <input
              type="text"
              placeholder="Search posts, authors, notes, tags..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-10 pr-4 py-2 text-sm bg-slate-100 dark:bg-slate-800 border-none rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
            />
          </div>

          <div className="flex items-center gap-3">
            <select
              value={sortBy}
              onChange={(e) => setSortBy(e.target.value)}
              className="text-xs bg-slate-100 dark:bg-slate-800 px-3 py-2 rounded-lg border-none outline-none font-medium text-slate-600 dark:text-slate-300 cursor-pointer"
            >
              <option value="newest">Sort: Newest First</option>
              <option value="oldest">Sort: Oldest First</option>
              <option value="author">Sort: Author A-Z</option>
            </select>

            <button
              onClick={() => setActiveModal('addManual')}
              className="flex items-center gap-1.5 px-3 py-2 rounded-lg bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold shadow-sm transition"
            >
              <Plus size={16} /> Add Post
            </button>
          </div>
        </header>

        {/* Content Feed */}
        <div className="flex-1 overflow-y-auto p-6 bg-slate-50 dark:bg-slate-950">
          <div className="max-w-6xl mx-auto space-y-4">
            <div className="flex items-center justify-between pb-2">
              <h2 className="text-xl font-bold text-slate-800 dark:text-white flex items-center gap-2">
                <span>{selectedCategory}</span>
                <span className="text-sm font-normal text-slate-500">({filteredPosts.length} posts found)</span>
              </h2>
            </div>

            {/* Empty State */}
            {filteredPosts.length === 0 && (
              <div className="text-center py-16 border-2 border-dashed border-slate-200 dark:border-slate-800 rounded-xl bg-white dark:bg-slate-900 p-8">
                <BookOpen size={40} className="mx-auto text-slate-400 mb-3" />
                <h3 className="text-lg font-bold text-slate-700 dark:text-slate-300">No posts found</h3>
                <p className="text-sm text-slate-500 max-w-sm mx-auto mt-1 mb-4">
                  Import your LinkedIn spreadsheet or click "Add Post" to start building your personal library.
                </p>
                <button
                  onClick={() => setActiveModal('import')}
                  className="px-4 py-2 bg-blue-600 text-white font-medium rounded-lg text-sm hover:bg-blue-700 inline-flex items-center gap-2"
                >
                  <Upload size={16} /> Import Spreadsheet
                </button>
              </div>
            )}

            {/* Cards Grid */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {filteredPosts.map((post) => (
                <div
                  key={post.id}
                  onClick={() => setSelectedPostDetail(post)}
                  className="group relative bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-4 flex flex-col justify-between hover:shadow-md hover:border-blue-400 dark:hover:border-blue-500 transition cursor-pointer"
                >
                  <div>
                    <div className="flex items-start justify-between gap-2 mb-2">
                      <span className="text-xs font-semibold px-2 py-0.5 rounded bg-blue-50 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400">
                        {post.category}
                      </span>
                      <div className="flex items-center gap-1">
                        <button
                          onClick={(e) => toggleFavorite(post, e)}
                          className={`p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 transition ${
                            post.favorite ? 'text-amber-500' : 'text-slate-400'
                          }`}
                        >
                          <Bookmark size={16} fill={post.favorite ? 'currentColor' : 'none'} />
                        </button>
                        <button
                          onClick={(e) => cycleStatus(post, e)}
                          className="text-xs px-2 py-0.5 rounded font-medium bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300"
                        >
                          {post.readStatus || 'unread'}
                        </button>
                      </div>
                    </div>

                    <h3 className="font-semibold text-sm text-slate-900 dark:text-white line-clamp-1">{post.author}</h3>
                    <p className="text-xs text-slate-500 dark:text-slate-400 mt-1 line-clamp-4 leading-relaxed whitespace-pre-wrap">
                      {post.text}
                    </p>

                    {post.tags && post.tags.length > 0 && (
                      <div className="flex flex-wrap gap-1 mt-3">
                        {post.tags.slice(0, 3).map((tag, i) => (
                          <span key={i} className="text-[10px] text-slate-500 dark:text-slate-400 bg-slate-100 dark:bg-slate-800 px-1.5 py-0.5 rounded">
                            #{tag}
                          </span>
                        ))}
                      </div>
                    )}
                  </div>

                  <div className="mt-4 pt-3 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between text-xs">
                    {post.postLink ? (
                      <a
                        href={post.postLink}
                        target="_blank"
                        rel="noopener noreferrer"
                        onClick={(e) => e.stopPropagation()}
                        className="text-blue-600 dark:text-blue-400 flex items-center gap-1 font-medium hover:underline"
                      >
                        LinkedIn <ExternalLink size={12} />
                      </a>
                    ) : (
                      <span className="text-slate-400">Manual Entry</span>
                    )}

                    <div className="flex items-center gap-2">
                      {post.notes && <span className="text-[10px] bg-emerald-50 dark:bg-emerald-950/40 text-emerald-600 px-1.5 py-0.5 rounded font-medium">Notes</span>}
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setEditingPost(post);
                        }}
                        className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200"
                      >
                        <Edit3 size={14} />
                      </button>
                      <button
                        onClick={(e) => handleDeletePost(post.id, e)}
                        className="text-slate-400 hover:text-red-500"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </main>

      {/* Post Detail Drawer */}
      {selectedPostDetail && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex justify-end">
          <div className="w-full max-w-xl bg-white dark:bg-slate-900 h-full p-6 overflow-y-auto flex flex-col justify-between shadow-2xl">
            <div>
              <div className="flex items-center justify-between pb-4 border-b border-slate-200 dark:border-slate-800">
                <span className="text-xs font-semibold px-2 py-1 bg-blue-100 text-blue-700 dark:bg-blue-900 dark:text-blue-200 rounded">
                  {selectedPostDetail.category}
                </span>
                <button onClick={() => setSelectedPostDetail(null)} className="text-slate-400 hover:text-slate-600">
                  <X size={20} />
                </button>
              </div>

              <div className="mt-4">
                <div className="flex items-center justify-between">
                  <h2 className="text-lg font-bold text-slate-900 dark:text-white">{selectedPostDetail.author}</h2>
                  {selectedPostDetail.postLink && (
                    <a
                      href={selectedPostDetail.postLink}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-blue-600 text-white rounded-md text-xs font-medium hover:bg-blue-700"
                    >
                      Open in LinkedIn <ExternalLink size={12} />
                    </a>
                  )}
                </div>

                <div className="mt-4 p-4 rounded-lg bg-slate-50 dark:bg-slate-800 text-sm whitespace-pre-wrap leading-relaxed text-slate-800 dark:text-slate-200 max-h-72 overflow-y-auto">
                  {selectedPostDetail.text}
                </div>

                {selectedPostDetail.extractedLinks?.length > 0 && (
                  <div className="mt-4">
                    <p className="text-xs font-semibold uppercase text-slate-400 mb-2">Attached Resources</p>
                    <div className="space-y-1">
                      {selectedPostDetail.extractedLinks.map((link, idx) => (
                        <a
                          key={idx}
                          href={link}
                          target="_blank"
                          rel="noreferrer"
                          className="block text-xs text-blue-500 hover:underline truncate"
                        >
                          🔗 {link}
                        </a>
                      ))}
                    </div>
                  </div>
                )}

                <div className="mt-6">
                  <label className="text-xs font-semibold uppercase text-slate-400 mb-2 block">Personal Notes</label>
                  <textarea
                    rows={4}
                    value={selectedPostDetail.notes || ''}
                    onChange={(e) => {
                      const updated = { ...selectedPostDetail, notes: e.target.value };
                      handleUpdatePost(updated);
                    }}
                    placeholder="Write key technical points, interview questions, or your implementation notes here..."
                    className="w-full text-sm p-3 rounded-lg border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 outline-none focus:ring-2 focus:ring-blue-500"
                  />
                </div>
              </div>
            </div>

            <div className="mt-6 pt-4 border-t border-slate-200 dark:border-slate-800 flex justify-end">
              <button
                onClick={() => setSelectedPostDetail(null)}
                className="px-4 py-2 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 text-xs font-semibold rounded-md"
              >
                Done
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Import Modal */}
      {activeModal === 'import' && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-900 rounded-xl max-w-md w-full p-6 shadow-xl relative border border-slate-200 dark:border-slate-800">
            <button
              onClick={() => {
                setActiveModal(null);
                setImportStats(null);
              }}
              className="absolute top-4 right-4 text-slate-400 hover:text-slate-600"
            >
              <X size={20} />
            </button>

            <h3 className="text-lg font-bold mb-1">Import Saved Posts</h3>
            <p className="text-xs text-slate-500 mb-4">
              Upload your LinkedIn export (.xlsx, .csv, or .json). Existing records will preserve personal notes.
            </p>

            <input
              type="file"
              ref={fileInputRef}
              accept=".xlsx, .xls, .csv, .json"
              onChange={handleFileUpload}
              className="block w-full text-sm text-slate-500 file:mr-4 file:py-2.5 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 dark:file:bg-slate-800 dark:file:text-blue-400 cursor-pointer"
            />

            {importStats && (
              <div className="mt-4 p-3 bg-emerald-50 dark:bg-emerald-950/40 border border-emerald-200 dark:border-emerald-800 rounded-lg text-xs space-y-1">
                <p className="font-bold text-emerald-800 dark:text-emerald-300">Import Successful!</p>
                <p className="text-emerald-700 dark:text-emerald-400">Total rows processed: {importStats.total}</p>
                <p className="text-emerald-700 dark:text-emerald-400">New posts added: {importStats.added}</p>
                <p className="text-emerald-700 dark:text-emerald-400">Existing posts updated: {importStats.updated}</p>
              </div>
            )}
          </div>
        </div>
      )}

      {/* Edit / Manual Add Modal */}
      {(editingPost || activeModal === 'addManual') && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-900 rounded-xl max-w-lg w-full p-6 shadow-xl border border-slate-200 dark:border-slate-800">
            <h3 className="text-lg font-bold mb-4">{editingPost ? 'Edit Post' : 'Add Post Manually'}</h3>

            <div className="space-y-3">
              <div>
                <label className="text-xs font-semibold text-slate-500">Author Name</label>
                <input
                  type="text"
                  defaultValue={editingPost?.author || ''}
                  id="authorInput"
                  className="w-full mt-1 px-3 py-2 text-sm bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg outline-none"
                />
              </div>

              <div>
                <label className="text-xs font-semibold text-slate-500">LinkedIn Post URL</label>
                <input
                  type="url"
                  defaultValue={editingPost?.postLink || ''}
                  id="linkInput"
                  className="w-full mt-1 px-3 py-2 text-sm bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg outline-none"
                />
              </div>

              <div>
                <label className="text-xs font-semibold text-slate-500">Category</label>
                <select
                  defaultValue={editingPost?.category || 'Oracle HCM'}
                  id="catInput"
                  className="w-full mt-1 px-3 py-2 text-sm bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg outline-none"
                >
                  {CATEGORIES.filter((c) => c !== 'All').map((c) => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="text-xs font-semibold text-slate-500">Post Text</label>
                <textarea
                  rows={4}
                  defaultValue={editingPost?.text || ''}
                  id="textInput"
                  className="w-full mt-1 px-3 py-2 text-sm bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg outline-none"
                />
              </div>
            </div>

            <div className="mt-5 flex justify-end gap-2">
              <button
                onClick={() => {
                  setEditingPost(null);
                  setActiveModal(null);
                }}
                className="px-4 py-2 text-xs font-semibold rounded-lg bg-slate-100 dark:bg-slate-800 hover:bg-slate-200"
              >
                Cancel
              </button>
              <button
                onClick={async () => {
                  const author = document.getElementById('authorInput').value;
                  const postLink = document.getElementById('linkInput').value;
                  const category = document.getElementById('catInput').value;
                  const text = document.getElementById('textInput').value;

                  if (editingPost) {
                    await handleUpdatePost({ ...editingPost, author, postLink, category, text });
                  } else {
                    const newPost = {
                      id: `post-${Date.now()}`,
                      postId: `${Date.now()}`,
                      author: author || 'Unknown',
                      postLink,
                      category,
                      text,
                      tags: (text.match(/#\w+/g) || []).map((t) => t.substring(1)),
                      favorite: false,
                      readStatus: 'unread',
                      notes: '',
                      summary: '',
                      importedAt: new Date().toISOString()
                    };
                    const nextList = [newPost, ...posts];
                    setPosts(nextList);
                    await savePostsToDB([newPost]);
                  }
                  setEditingPost(null);
                  setActiveModal(null);
                }}
                className="px-4 py-2 text-xs font-semibold rounded-lg bg-blue-600 text-white hover:bg-blue-700"
              >
                Save
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}