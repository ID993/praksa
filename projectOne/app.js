const express = require('express');
const app = express();
const bodyParser = require('body-parser');
const data = require('./data.json');
const fs = require('fs');

const currentDate = new Date();
const options = {
  year: 'numeric',
  month: '2-digit',
  day: '2-digit',
  hour: '2-digit',
  minute: '2-digit',
  second: '2-digit',
};

app.use(bodyParser.json());


const port = 3000;

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});

app.get('/users/:userID', (req, res) => {
    const userID = req.params.userID;
    const user = data.users.find(user => user.id === parseInt(userID));
    
    if (user) {
      res.json(user);
    } else {
      res.status(404).json({ error: 'User not found' });
    }
  });
  
app.get('/posts/:postID', (req, res) => {
    const postID = req.params.postID;
    const post = data.posts.find(post => post.id === parseInt(postID));

    if (post) {
        res.json(post);
    } else {
        res.status(404).json({ error: 'Post not found' });
    }
});

app.get('/posts', (req, res) => {
    const { DatumOd, DatumDo } = req.query;
    const startDate = new Date(DatumOd);
    const endDate = new Date(DatumDo);
    
    const posts = data.posts.filter(post => {
      const postDate = new Date(post.last_update);
      return postDate >= startDate && postDate <= endDate;
    });
    
    res.json(posts);
  });

  app.post('/users/:userID', (req, res) => {
    const userID = req.params.userID;
    const { noviEmail } = req.body;
  
    const user = data.users.find(user => user.id === parseInt(userID));
  
    if (user) {
      user.email = noviEmail;
  
      fs.writeFile('./data.json', JSON.stringify(data, null, 2), err => {
        if (err) {
          res.status(500).json({ error: 'Unable to update user email' });
        } else {
          res.json(user);
        }
      });
    } else {
      res.status(404).json({ error: 'User not found' });
    }
  });

app.put('/posts', (req, res) => {
    const { userID, title, body } = req.body;
    
    const newPost = {
      id: data.posts.length + 1,
      title,
      body,
      user_id: userID,
      last_update: currentDate.toLocaleDateString(options)
    };
    
    data.posts.push(newPost);
    
    fs.writeFile('./data.json', JSON.stringify(data, null, 2), err => {
        if (err) {
          res.status(500).json({ error: 'Unable to create new post' });
        } else {
          res.json(newPost);
        }
      });
  });
  
  
app.get('/', (req, res) => {
    res.send('Hello, World!');
  });
  
