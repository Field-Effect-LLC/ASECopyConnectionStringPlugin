using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aga.Controls.Tree;
using MsaSQLEditor.DockWindows;
using MsaSQLEditor.Interfaces;
using TreeView = MsaSQLEditor.DockWindows.TreeView;

namespace CopyConnectionStringPlugin
{
    public class CopyConnectionStringPlugin : IPlugin2
    {
        IPluginContext2 _PluginContext;
        TreeView _QueriesWindow;
        TreeView _TablesWindow;
        ContextMenuStrip _QueriesContextMenu;
        ContextMenuStrip _TablesContextMenu;

        public string Name => "Copy Connection String Plugin";

        public void Initialize(IPluginContext2 context)
        {
            _PluginContext = context;

            _QueriesWindow = _PluginContext.SystemWindows
                .Where<ISystemWindow>(w => w.Text == "Queries")
                .Select<ISystemWindow, TreeView>(item => (TreeView)item)
                .FirstOrDefault();

            _TablesWindow = _PluginContext.SystemWindows
                .Where<ISystemWindow>(w => w.Text == "Tables")
                .Select<ISystemWindow, TreeView>(item => (TreeView)item)
                .FirstOrDefault();

            var treeViewType = typeof(TreeView);
            var propertyInfo = treeViewType.GetField("ItemsContextMenu", BindingFlags.Instance | BindingFlags.NonPublic);
            _QueriesContextMenu = (ContextMenuStrip)propertyInfo.GetValue(_QueriesWindow);
            _TablesContextMenu = (ContextMenuStrip)propertyInfo.GetValue(_TablesWindow);

            var queriesMenuItem = AddContextMenuItem(_QueriesContextMenu);
            var tablesMenuItem = AddContextMenuItem(_TablesContextMenu);

            CancelEventHandler contextMenuOpening = (s, e) =>
            {
                ContextMenuStrip menuStrip = (ContextMenuStrip)s;
                TreeView srcView = GetFormFromMenu(menuStrip);

                string connectionString = SelectedNodeConnectionString(srcView);
                queriesMenuItem.Enabled = !String.IsNullOrWhiteSpace(connectionString);
                tablesMenuItem.Enabled = !String.IsNullOrWhiteSpace(connectionString);
            };

            _QueriesContextMenu.Opening += contextMenuOpening;
            _TablesContextMenu.Opening += contextMenuOpening;
            
        }

        private TreeView GetFormFromMenu(ContextMenuStrip menuStrip)
        {
            if (menuStrip.Equals(_QueriesContextMenu))
                return _QueriesWindow;

            if (menuStrip.Equals(_TablesContextMenu))
                return _TablesWindow;

            return null;
        }

        private string SelectedNodeConnectionString(TreeView view)
        {
            var nodes = view.SelectedNodes.ToList<TreeNodeAdv>();
            string connectionString = String.Empty;

            if (nodes.Count != 1)
                return connectionString;

            TreeNodeAdv node = nodes[0];
            
            if (view.Equals(_TablesWindow))
            {
                connectionString = (_PluginContext.Database as dynamic)
                    .Database.TableDefs[node.ToString()].Connect;
            }

            if (view.Equals(_QueriesWindow))
            {
                connectionString = (_PluginContext.Database as dynamic)
                    .Database.QueryDefs[node.ToString()].Connect;
            }

            return connectionString;
        }

        private ToolStripMenuItem AddContextMenuItem(ContextMenuStrip itemsContextMenu)
        {
            ToolStripMenuItem item = new ToolStripMenuItem("Copy Connection String to Clipboard");
            item.Name = "CopyConnxString";
            item.Click += HandleMenuItem;
            item.Tag = true; //Disable on multiple selection
            itemsContextMenu.Items.Add(item);
            return item;
        }

        private void HandleMenuItem(object sender, EventArgs arg)
        {
            ToolStripMenuItem item = (ToolStripMenuItem)sender;
            var contextMenuStrip = (ContextMenuStrip)item.GetCurrentParent();
            TreeView srcView = GetFormFromMenu(contextMenuStrip);
            string connectionString = SelectedNodeConnectionString(srcView);
            Clipboard.SetText(connectionString);
        }

        public void Unload()
        {
            _PluginContext = null;
        }
    }
}
