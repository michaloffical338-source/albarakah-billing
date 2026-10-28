if st.sidebar.button("🚚 Migrate Old Auth"):
    if os.path.exists("users.json"):
        with open("users.json", "r", encoding="utf-8") as f:
            old_users = json.load(f)
        cloud_users = load_users()
        for u, info in old_users.items():
            cloud_users[u] = info
        save_users(cloud_users)
        st.sidebar.success(f"✅ {len(old_users)} users migrated")
    if os.path.exists("bookers_registry.json"):
        with open("bookers_registry.json", "r", encoding="utf-8") as f:
            old_bk = json.load(f)
        cloud_bk = load_bookers()
        for u, info in old_bk.items():
            cloud_bk[u] = info
        save_bookers(cloud_bk)
        st.sidebar.success(f"✅ {len(old_bk)} bookers migrated")
