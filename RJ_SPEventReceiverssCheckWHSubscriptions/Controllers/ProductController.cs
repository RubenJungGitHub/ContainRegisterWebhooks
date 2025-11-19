using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Linq;

namespace RJ_SPEventReceiversASPWebApp.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ProductsController : ControllerBase
    {
        private static readonly List<Product> _products = new()
        {
            new Product { Id = 1, Name = "Laptop", Price = 1500 },
            new Product { Id = 2, Name = "Mouse", Price = 25 }
        };

        // GET api/products
        [HttpGet]
        public IActionResult GetAll()
        {
            return Ok(_products);
        }

        // GET api/products/2
        [HttpGet("{id}")]
        public IActionResult GetById(int id)
        {
            var product = _products.FirstOrDefault(p => p.Id == id);
            if (product == null)
                return NotFound();
            return Ok(product);
        }

        // POST api/products
        [HttpPost]
        public IActionResult Create([FromBody] Product newProduct)
        {
            if (newProduct == null)
                return BadRequest("Invalid product");

            newProduct.Id = _products.Max(p => p.Id) + 1;
            _products.Add(newProduct);
            return CreatedAtAction(nameof(GetById), new { id = newProduct.Id }, newProduct);
        }

        // PUT api/products/2
        [HttpPut("{id}")]
        public IActionResult Update(int id, [FromBody] Product updatedProduct)
        {
            var existing = _products.FirstOrDefault(p => p.Id == id);
            if (existing == null)
                return NotFound();

            existing.Name = updatedProduct.Name;
            existing.Price = updatedProduct.Price;
            return NoContent();
        }

        // DELETE api/products/2
        [HttpDelete("{id}")]
        public IActionResult Delete(int id)
        {
            var product = _products.FirstOrDefault(p => p.Id == id);
            if (product == null)
                return NotFound();

            _products.Remove(product);
            return NoContent();
        }
    }

    // Simple model class
    public class Product
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public decimal Price { get; set; }
    }
}