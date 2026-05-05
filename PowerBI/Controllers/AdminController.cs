using Microsoft.AspNetCore.Mvc;
using PowerBI.Models;
using PowerBI.Data;
using Microsoft.EntityFrameworkCore;

namespace PowerBI.Controllers
{
    public class AdminController : Controller
    {
        private readonly AppDbContext _db;

        public AdminController(AppDbContext db)
        {
            _db = db;
        }

        public async Task<IActionResult> Index()
        {
            Console.WriteLine("[ADMIN] Loading Admin Dashboard Overview...");
            
            var users = await _db.Users.ToListAsync();
            var workspaces = await _db.Workspaces.ToListAsync();
            var reports = await _db.Reports.ToListAsync();

            ViewBag.UserCount = users.Count;
            ViewBag.WorkspaceCount = workspaces.Count;
            ViewBag.ReportCount = reports.Count;

            Console.WriteLine($"[ADMIN] System Stats - Users: {users.Count}, Workspaces: {workspaces.Count}, Reports: {reports.Count}");
            return View(users);
        }

        [HttpPost]
        public async Task<IActionResult> CreateUser(User user)
        {
            Console.WriteLine($"[ADMIN] Creating new user: {user.Email} (Role: {user.Role})");
            user.Role = user.Role ?? "User";
            _db.Users.Add(user);
            await _db.SaveChangesAsync();
            Console.WriteLine($"[ADMIN] User created with ID: {user.Id}");
            return RedirectToAction("Index");
        }

        public async Task<IActionResult> DeleteUser(int id)
        {
            Console.WriteLine($"[ADMIN] Deleting user ID: {id}");
            var user = await _db.Users.FindAsync(id);
            if (user != null)
            {
                _db.Users.Remove(user);
                await _db.SaveChangesAsync();
                Console.WriteLine("[ADMIN] User deleted successfully.");
            }
            return RedirectToAction("Index");
        }
    }
}
